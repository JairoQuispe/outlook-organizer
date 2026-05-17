# Guía completa de flags para `outlook-import-pst.ps1`

> Script: `core/scripts-pwsh/outlook-import-pst.ps1`
>
> Uso general:
> ```powershell
> pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 [flags]
> ```
>
> Requisitos: Outlook Desktop instalado (COM Automation), PowerShell 5.1+ o pwsh 7+, y un archivo PST accesible en disco local.

---

## Resumen del flujo de ejecución

1. **Valida parámetros** (PstPath, TargetStoreId, filtros de fecha/meses).
2. **Monta el PST** vía `Namespace.AddStoreEx` si no estaba previamente montado.
3. **Analiza contenido** — cuenta ítems/carpetas (o usa el plan si se proporcionó `-FolderPlanPath`).
4. **Importa ítems** — copia o mueve cada ítem al buzón destino con rate limiting (token bucket).
5. **Desmonta el PST** si fue montado por el script.
6. **Emite resultado** — un payload `restoreResult` con totales, fallos y tiempos.

---

## Parámetros obligatorios principales

| Flag | Tipo | Default | Descripción |
|------|------|---------|-------------|
| `-PstPath` | `string` | — | Ruta completa al archivo `.pst`. Debe existir en disco. El script lo monta automáticamente si no estaba ya abierto en Outlook. |
| `-TargetStoreId` | `string` | — | StoreID del buzón destino. Obtén los IDs ejecutando `outlook-list-stores.ps1 -Json`. |
| `-Action` | `Copy` \| `Move` | `Copy` | **Copy**: clona el ítem y lo mueve al destino (el original permanece en el PST). **Move**: mueve directamente — el ítem desaparece del PST. |

### Cuándo usar cada Action

| Escenario | Action recomendada | Razón |
|-----------|-------------------|-------|
| Migración de archivo histórico al buzón principal | `Copy` | Preserva el PST como respaldo intacto. |
| Consolidación de PSTs en un solo buzón (eliminar originales) | `Move` | Evita duplicados y reduce tamaño del PST fuente. |
| Pruebas o importación parcial | `Copy` | Permite re-ejecutar sin pérdida de datos. |

---

## Filtros de contenido

### `-FilterOnlyYear`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Rango | 1900–9999 |
| Default | Sin filtro (todos los años) |

Importa únicamente ítems cuyo `ReceivedTime` o `CreationTime` cae en el año especificado. Si el ítem no tiene fecha, se omite.

**Cuándo usarlo:**
- Migrar solo un año fiscal específico.
- Importar correos antiguos por etapas (año por año) para no saturar el buzón.

```powershell
-FilterOnlyYear 2022
```

---

### `-FilterOnlyMonths`

| Atributo | Valor |
|----------|-------|
| Tipo | `string` (lista separada por comas) |
| Acepta | Números (1–12), nombres completos o abreviaturas en español/inglés |
| Default | Sin filtro (todos los meses) |

Se puede combinar con `-FilterOnlyYear` para una ventana muy precisa.

**Tokens válidos (ejemplos):**
`1`, `ene`, `enero`, `jan`, `january`, `feb`, `febrero`, `february`, `mar`, `marzo`, `march`, `abr`, `abril`, `april`, `may`, `mayo`, `jun`, `junio`, `june`, `jul`, `julio`, `july`, `ago`, `agosto`, `aug`, `august`, `sep`, `set`, `septiembre`, `setiembre`, `september`, `oct`, `octubre`, `october`, `nov`, `noviembre`, `november`, `dic`, `diciembre`, `dec`, `december`.

**Cuándo usarlo:**
- Importar solo Q1 (enero–marzo) de un año.
- Migrar mes a mes para controlar el impacto en la cuota del buzón.

```powershell
-FilterOnlyYear 2023 -FilterOnlyMonths "ene,feb,mar"
-FilterOnlyMonths "1,2,3"
-FilterOnlyMonths "january,february,march"
```

---

### `-IncludeFolders`

| Atributo | Valor |
|----------|-------|
| Tipo | `string[]` |
| Default | Vacío (todas las carpetas) |

Lista de rutas relativas dentro del PST. Acepta separadores `/` o `\`. Incluye la carpeta indicada y todas sus subcarpetas.

**Cuándo usarlo:**
- Importar solo carpetas específicas sin tocar el resto del PST.
- Ejecución desde línea de comandos sin necesidad de JSON.

```powershell
-IncludeFolders "Bandeja de entrada\Clientes","Archivo\2024"
```

---

### `-IncludeFoldersJson`

| Atributo | Valor |
|----------|-------|
| Tipo | `string` (JSON array serializado) |
| Default | Vacío |

Equivalente a `-IncludeFolders` pero en formato JSON, pensado para integraciones programáticas (desde Bun, Node.js, etc.).

```powershell
-IncludeFoldersJson '["Bandeja de entrada/Clientes","Archivo/2024"]'
```

> **Nota:** Si se usan ambos (`-IncludeFolders` y `-IncludeFoldersJson`), se combinan en un solo conjunto unificado.

---

### `-FolderPlanPath`

| Atributo | Valor |
|----------|-------|
| Tipo | `string` (ruta a archivo JSON) |
| Default | Sin plan |

Acepta un archivo JSON generado por `outlook-scan-pst.ps1 -ExportFolders`. Estructura esperada:

```json
{
  "type": "folderExport",
  "folders": [
    { "path": "Bandeja de entrada\\Clientes", "itemCount": 340 },
    { "path": "Archivo\\2023", "itemCount": 1200 }
  ]
}
```

**Comportamiento:**
- El importador recorre las carpetas **exactamente en el orden** del JSON.
- Usa `itemCount` para calcular progreso real y ETA.
- Omite recursión en subcarpetas (cada carpeta del plan se procesa de forma independiente).
- Se fusiona con `-IncludeFolders`/`-IncludeFoldersJson` si están presentes.

**Cuándo usarlo:**
- Importaciones largas donde se necesita seguimiento de progreso preciso.
- Automatización con UI que consume eventos JSON para barras de progreso.
- Re-ejecuciones parciales (editar el JSON para quitar carpetas ya importadas).

```powershell
-FolderPlanPath "scan-year-2023-pst-archivo-with-size-20250512-183000.json"
```

---

### `-SkipDuplicates`

| Atributo | Valor |
|----------|-------|
| Tipo | `switch` |
| Default | Desactivado |

Antes de copiar/mover cada ítem, construye un índice de claves únicas en la carpeta destino y compara. Estrategia de clave (en orden de prioridad):

1. **Message-ID** (`mid:<normalized-id>`) — más confiable.
2. **SearchKey** MAPI (`sk:<hex>`) — backup binario.
3. **Clave compuesta** (`comp:<subject>|<date>|<sender>`) — fallback cuando no hay Message-ID ni SearchKey.

**Cuándo usarlo:**
- Re-ejecuciones de importación donde parte de los ítems ya fue migrada.
- Consolidación de múltiples PSTs que pueden tener correos superpuestos.

---

### `-DeepDuplicateCheck`

| Atributo | Valor |
|----------|-------|
| Tipo | `switch` |
| Default | Desactivado |

Solo tiene efecto cuando se combina con `-SkipDuplicates`. Indexa no solo la carpeta destino directa, sino también **todas sus subcarpetas**, lo que detecta duplicados que fueron movidos manualmente a subcarpetas.

**Cuándo usarlo:**
- Buzón destino donde el usuario reorganiza correos en subcarpetas.
- PSTs con estructura de carpetas diferente a la del buzón destino.

> **Impacto en rendimiento:** Incrementa significativamente el tiempo de indexación inicial. Usar solo si es necesario.

---

## Control de rendimiento y reintentos

El script implementa un **token bucket** para controlar la tasa de operaciones y un mecanismo de **backoff exponencial** para manejar throttling de Exchange/MAPI.

### `-ItemsPerMinute`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Default | `120` |

Tasa promedio de ítems procesados por minuto. El token bucket se rellena a `ItemsPerMinute / 60` tokens por segundo.

**Guía de valores:**

| Valor | Escenario |
|-------|-----------|
| 60–80 | Exchange Online (Office 365) — política de throttling estricta |
| 100–150 | Exchange on-premises con carga moderada |
| 200–500 | Servidor Exchange local dedicado o buzón PST-a-PST |

---

### `-BurstSize`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Default | `20` |

Capacidad máxima del token bucket. Permite procesar una ráfaga de ítems consecutivos antes de frenar.

**Cuándo ajustarlo:**
- Aumentar (30–50) si el servidor tolera ráfagas y quieres acelerar carpetas pequeñas.
- Reducir (5–10) si recibes errores de throttling frecuentes.

---

### `-AdaptiveThrottling`

| Atributo | Valor |
|----------|-------|
| Tipo | `switch` |
| Default | Desactivado |

Cuando está activo:
- Al detectar un error de throttling, **reduce el multiplicador** de tasa (×0.7 normal, ×0.5 si es desconexión de red).
- Tras 20 operaciones exitosas consecutivas, **recupera** el multiplicador gradualmente (×1.15).
- La tasa nunca baja de 10% del original ni sube por encima de 100%.

**Cuándo usarlo:**
- Importaciones contra Exchange Online donde la política de throttling es dinámica.
- Ejecuciones nocturnas/desatendidas donde no puedes supervisar errores.

```powershell
-AdaptiveThrottling -ItemsPerMinute 120 -BurstSize 20
```

---

### `-MaxRetries`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Default | `5` |

Número máximo de reintentos para una operación COM que falla por throttling o desconexión. Después de agotar los reintentos, el ítem se marca como fallido.

---

### `-InitialBackoffMs`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` (milisegundos) |
| Default | `1000` |

Tiempo de espera base antes del primer reintento. Se duplica en cada intento sucesivo (backoff exponencial): 1s → 2s → 4s → 8s → 16s...

---

### `-MaxBackoffMs`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` (milisegundos) |
| Default | `30000` (30 segundos) |

Límite superior del backoff exponencial. Ningún reintento esperará más de este valor.

> **Para errores de red/MAPI** el script aplica un mínimo de 5 segundos y un multiplicador de ×1.5 (hasta 60s) para dar tiempo a la reconexión.

---

### `-DuplicateIndexInactivityTimeoutSec`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Rango | 0–7200 |
| Default | `180` (3 minutos) |

Si durante la construcción del índice de duplicados no hay progreso (ej: MAPI colgado) por este tiempo, se aborta la indexación y se continúa con un índice parcial.

**Cuándo ajustarlo:**
- Poner `0` para deshabilitarlo (sin timeout) en redes rápidas.
- Aumentar a 300–600 en carpetas destino con >50,000 ítems en redes lentas.

---

## Manejo de tamaño y fallos

### `-MaxItemSizeBytes`

| Atributo | Valor |
|----------|-------|
| Tipo | `long` |
| Default | `157286400` (150 MB) |

Ítems que superan este tamaño se registran como `too_large` en los fallos y se saltan.

| Valor | Descripción |
|-------|-------------|
| `157286400` | 150 MB — default conservador |
| `52428800` | 50 MB — para buzones con cuota limitada |
| `0` | Sin límite — procesar todo sin importar tamaño |

---

### `-MaxFailureRecords`

| Atributo | Valor |
|----------|-------|
| Tipo | `int` |
| Rango | 0–100000 |
| Default | `1000` |

Número máximo de fallos detallados que se almacenan en el resultado. Los fallos adicionales se contabilizan en `failureOverflow` pero no se almacenan individualmente (evita consumo excesivo de memoria).

---

## Modo headless / JSON

### `-Json`

| Atributo | Valor |
|----------|-------|
| Tipo | `switch` |
| Default | Desactivado |

Convierte **toda la salida** (logs, progreso, errores, resultado) en líneas JSON compactas escritas a stdout con encoding UTF-8. Tipos de eventos emitidos:

| `type` | Descripción |
|--------|-------------|
| `log` | Mensaje de log con `level`, `message`, `timestamp` |
| `progress` | Barra de progreso con `percent`, `copied`, `moved`, `skipped`, `failed` |
| `error` | Error fatal |
| `throttleStats` | Métricas del token bucket (cada 10s) |
| `dupSkipped` | Ítem duplicado omitido (con `folder`, `subject`, `key`, `source`) |
| `restoreResult` | Resultado final completo |

**Cuándo usarlo:**
- Integración con frontend (Bun, Node.js, Electron, OpenTUI).
- Parsing automatizado de resultados.

---

### `-Headless`

| Atributo | Valor |
|----------|-------|
| Tipo | `switch` |
| Default | Desactivado |

Funcionalmente idéntico a `-Json` en cuanto a silenciar `Write-Host` y `Write-Progress`. Se puede usar solo si quieres una ejecución silenciosa sin formato JSON.

> **Tip:** `-Json` implica `-Headless`. No es necesario pasarlos juntos, aunque es común hacerlo para claridad.

---

## Ejemplos completos

### 1. Importación básica (interactiva)

Copia todos los correos de un PST al buzón destino con barra de progreso en terminal.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo-2023.pst" `
  -TargetStoreId "00000000C..." `
  -Action Copy
```

---

### 2. Importar solo un año y meses específicos

Migrar solo Q1 2023 (enero a marzo) desde un PST grande.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\historico.pst" `
  -TargetStoreId "00000000C..." `
  -FilterOnlyYear 2023 `
  -FilterOnlyMonths "ene,feb,mar" `
  -SkipDuplicates
```

---

### 3. Importar carpetas específicas con detección profunda de duplicados

Solo las carpetas "Bandeja de entrada\Clientes" y "Archivo\Contratos", detectando duplicados incluso en subcarpetas del destino.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\empresa.pst" `
  -TargetStoreId "00000000C..." `
  -IncludeFolders "Bandeja de entrada\Clientes","Archivo\Contratos" `
  -SkipDuplicates -DeepDuplicateCheck
```

---

### 4. Modo headless con JSON para integración programática

Ejecución desatendida con carpetas en JSON, throttling adaptativo y salida JSON para consumir desde un proceso padre.

```powershell
$folders = '["Inbox/Clientes","Archivo/2023","Elementos enviados"]'
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -TargetStoreId "00000000C..." `
  -Json -Headless `
  -IncludeFoldersJson $folders `
  -FilterOnlyYear 2023 -FilterOnlyMonths "ene,feb" `
  -SkipDuplicates -DeepDuplicateCheck `
  -AdaptiveThrottling -ItemsPerMinute 80
```

---

### 5. Importar usando un plan de carpetas exportado

Usa el archivo generado por `outlook-scan-pst.ps1 -ExportFolders` para seguimiento preciso de progreso.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -TargetStoreId "00000000C..." `
  -FolderPlanPath "scan-year-2023-pst-archivo-with-size-20250512-183000.json" `
  -SkipDuplicates `
  -AdaptiveThrottling
```

- El importador recorre las carpetas exactamente en el orden definido en el JSON.
- Cada ítem muestra logs del estilo `carpeta Inbox/Clientes - item 320 de 640 - cargando 50%`, basados en `itemCount` exportado, útil para estimar ETA.

---

### 6. Mover ítems (eliminar del PST) con tasa conservadora

Migración definitiva con tasa baja para Exchange Online.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\legacy.pst" `
  -TargetStoreId "00000000C..." `
  -Action Move `
  -ItemsPerMinute 60 -BurstSize 10 `
  -MaxRetries 8 -MaxBackoffMs 45000 `
  -AdaptiveThrottling `
  -SkipDuplicates
```

---

### 7. Importar ítems grandes sin límite de tamaño

Desactivar el filtro de tamaño para importar adjuntos grandes (>150 MB).

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\adjuntos-grandes.pst" `
  -TargetStoreId "00000000C..." `
  -MaxItemSizeBytes 0 `
  -ItemsPerMinute 30 -BurstSize 5
```

> **Precaución:** Exchange Online tiene un límite de 150 MB por mensaje. Superar este límite resultará en errores del servidor aunque el script no filtre el ítem.

---

### 8. Re-ejecución segura tras interrupción

Tras una interrupción (error de red, reinicio), re-ejecutar sin duplicar lo ya importado.

```powershell
pwsh -File core/scripts-pwsh/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -TargetStoreId "00000000C..." `
  -SkipDuplicates -DeepDuplicateCheck `
  -DuplicateIndexInactivityTimeoutSec 300
```

---

## Mapeo automático de carpetas conocidas

El script mapea automáticamente las carpetas raíz del PST a las carpetas predeterminadas de Outlook:

| Carpeta fuente (PST) | Carpeta destino (OlDefaultFolders) |
|----------------------|-----------------------------------|
| `Bandeja de entrada` / `Inbox` | Bandeja de entrada (olFolderInbox = 6) |
| `Elementos eliminados` / `Deleted Items` | Elementos eliminados (olFolderDeletedItems = 3) |
| `Elementos enviados` / `Sent Items` | Elementos enviados (olFolderSentMail = 5) |
| `Borradores` / `Drafts` | Borradores (olFolderDrafts = 16) |
| `Correo no deseado` / `Junk Email` | Correo no deseado (olFolderJunk = 23) |
| `Bandeja de salida` / `Outbox` | Bandeja de salida (olFolderOutbox = 4) |
| Cualquier otra carpeta | Se crea como subcarpeta bajo la raíz del buzón destino |

---

## Errores detectados como throttling

El mecanismo de reintentos se activa ante estos patrones:

- **HRESULT:** `0x80040115`, `0x8004011D`, `0x80040600`, `0x800401FD`
- **Texto:** `Server Busy`, `throttl*`, `budget`, `too many requests`, `429`
- **Red/MAPI:** `problemas en la red`, `conexión con Microsoft Exchange`, `no se puede completar`, `network`, `disconnected`, `RPC_E_DISCONNECTED`

---

## Salida final (`restoreResult`)

Al completar, el script emite un payload con:

```json
{
  "type": "restoreResult",
  "filterOnlyYear": 2023,
  "filterOnlyMonths": [1, 2, 3],
  "copied": 1542,
  "moved": 0,
  "skipped": 87,
  "failed": 3,
  "elapsedMs": 324560,
  "throttleEvents": 2,
  "totalWaitedMs": 15000,
  "failures": [
    { "folder": "Inbox\\Clientes", "subject": "Contrato v2", "reason": "too_large", "sizeBytes": 200000000 }
  ],
  "failureOverflow": 0
}
```

En modo interactivo (sin `-Json`), se imprime un resumen tabular en consola.

---

## Listar carpetas antes de importar

Para explorar un PST sin ejecutar el importador (ya sea por ruta o por StoreId existente), usa el script dedicado:

```powershell
pwsh -File core/scripts-pwsh/outlook-scan-pst.ps1 `
  [-PstPath "C:\PSTs\archivo.pst" | -StoreId "00000000C..."] `
  -Json [-PreserveSession]
```

Emite eventos `scanMeta`, `scanProgress` y una lista plana de carpetas con conteos por año, idénticos a los que consume la UI de Bun/OpenTUI.

- Admite `-FilterOnlyYear` para limitar la ventana de carpetas que se incluyen en el listado (igual que en el importador) y `-IncludeSize` para agregar `sizeBytes`/`sizeHuman` a cada carpeta usando la columna `Size` de la tabla MAPI cuando está disponible.
- El flag `-Summary` emite un único `type: "summary"` JSON con metadatos, totales, desglose por año y top folders; `-ExportResult` escribe automáticamente ese payload (o la lista de carpetas si no pediste summary) a `scan-result-<timestamp>.json` (por defecto) o `.txt` si pasas `-ExportResult text`. La versión `text` usa saltos de línea para mostrar el mismo resumen de forma legible.

---

## Flujo recomendado completo

```
1. Listar stores      →  outlook-list-stores.ps1 -Json
2. Escanear PST       →  outlook-scan-pst.ps1 -PstPath ... -ExportFolders -Json
3. (Opcional) Revisar  →  Editar el JSON del plan para filtrar carpetas
4. Importar           →  outlook-import-pst.ps1 -PstPath ... -FolderPlanPath ... -SkipDuplicates
```

---

## Notas adicionales

- Para obtener `TargetStoreId`, ejecuta `core/scripts-pwsh/outlook-list-stores.ps1` (opcional `-Json`).
- Si Outlook ya está abierto, el script lo reutiliza; si no, lo inicia automáticamente.
- Al finalizar, si el PST fue montado por el script, se desmonta automáticamente.
- Ante interrupciones, valida que Outlook no quede en segundo plano antes de la siguiente ejecución.
- Los nombres de carpeta se normalizan (Unicode FormKC, remoción de chars invisibles, case-insensitive) para matching robusto.
- Si una carpeta no se puede crear en el destino (permisos, límites), se marca como fallida y se omiten futuros intentos para esa misma carpeta.
