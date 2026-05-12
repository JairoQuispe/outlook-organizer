# Guía de flags para `outlook-import-pst.ps1`

> Script: `scripts-powershell/outlook-import-pst.ps1`
>
> Uso general:
> ```powershell
> pwsh -File scripts-powershell/outlook-import-pst.ps1 [flags]
> ```
>
> Requisitos: Outlook instalado, permisos de PowerShell para automatización y un PST accesible.

---

## Parámetros obligatorios principales

| Flag | Tipo | Descripción |
|------|------|-------------|
| `-PstPath` | `string` | Ruta completa al archivo PST que se montará. |
| `-TargetStoreId` | `string` | StoreID del buzón destino (obtén los IDs con `outlook-list-stores.ps1`). |
| `-Action` | `Copy`/`Move` (default `Copy`) | Controla si los ítems se copian o se mueven del PST al buzón. |

## Filtros de contenido

| Flag | Tipo | Notas |
|------|------|-------|
| `-FilterOnlyYear` | `int` (1900–9999) | Importa únicamente ítems cuyo año (Received/Creation) coincide. |
| `-FilterOnlyMonths` | `string` (ej. `"ene,feb,jun"`) | Acepta números o nombres/abreviaturas (en español o inglés). Se normaliza y se compara vía hashset. |
| `-IncludeFolders` | `string[]` | Lista simple de rutas relativas dentro del PST. Ej: `"Inbox/Clientes"`. |
| `-IncludeFoldersJson` | `string` (JSON array) | Versión serializada para integraciones. Ej: `'["Inbox/Clientes","Archivo\\2024"]'`. |
| `-SkipDuplicates` | `switch` | Construye un índice de duplicados (Message-ID, SearchKey o clave compuesta) y omite coincidencias Encontradas. |
| `-DeepDuplicateCheck` | `switch` | Cuando se usa con `-SkipDuplicates`, también indexa subcarpetas destino para claves existentes. |
| `-ListFolders` | `switch` | No importa datos; recorre el PST y emite metadatos + lista plana de carpetas (útil para explorar antes de importar). |

## Control de rendimiento y reintentos

| Flag | Tipo | Descripción |
|------|------|-------------|
| `-ItemsPerMinute` | `int` (default 120) | Tasa promedio objetivo para la cubeta de tokens. |
| `-BurstSize` | `int` (default 20) | Capacidad del bucket (ráfaga máxima). |
| `-AdaptiveThrottling` | `switch` | Reduce dinámicamente la tasa cuando detecta throttling o desconexiones. |
| `-MaxRetries` | `int` (default 5) | Intentos de reintento para operaciones COM con throttling. |
| `-InitialBackoffMs` | `int` (default 1000) | Base del backoff exponencial. |
| `-MaxBackoffMs` | `int` (default 30000) | Límite superior del backoff. |
| `-DuplicateIndexInactivityTimeoutSec` | `int` (0–7200, default 180) | Si no hay progreso al indexar duplicados durante este tiempo, se sale con un índice parcial. |

## Manejo de tamaño y fallos

| Flag | Tipo | Descripción |
|------|------|-------------|
| `-MaxItemSizeBytes` | `long` (default 157286400 ≈ 150 MB) | Ítems mayores se marcan como `too_large`, se registran y se saltan. Usa `0` para deshabilitar el límite. |
| `-MaxFailureRecords` | `int` (0–100000, default 1000) | Número máximo de fallos detallados que se guardan en la salida. El resto se cuenta en `failureOverflow`. |

## Modo headless / JSON

| Flag | Tipo | Descripción |
|------|------|-------------|
| `-Json` | `switch` | Convierte todos los logs/progresos/resultados en JSON compacto UTF-8 (pensado para Bun/OpenTUI). |
| `-Headless` | `switch` | Evita UI (idéntico a `-Json` en cuanto a silenciamiento de `Write-Host`; se puede usar sólo para CLI silenciosa). |

### Ejemplos

#### 1. Importación básica
```powershell
pwsh -File scripts-powershell/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -TargetStoreId "00000000C..." `
  -Action Copy
```

#### 2. Headless + duplicados profundos + filtros
```powershell
$folders = '["Inbox/Clientes","Archivo/2023"]'
pwsh -File scripts-powershell/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -TargetStoreId "00000000C..." `
  -Json -Headless `
  -IncludeFoldersJson $folders `
  -FilterOnlyYear 2023 -FilterOnlyMonths "ene,feb" `
  -SkipDuplicates -DeepDuplicateCheck
```

#### 3. Sólo listar carpetas
```powershell
pwsh -File scripts-powershell/outlook-import-pst.ps1 `
  -PstPath "C:\PSTs\archivo.pst" `
  -ListFolders -Json
```

### Notas adicionales
- Para obtener `TargetStoreId`, ejecuta `scripts-powershell/outlook-list-stores.ps1` (opcional `-Json`).
- Si Outlook ya está abierto, el script lo reutiliza; si no, lo inicia y al final desmonta el PST agregado.
- Ante interrupciones, valida que Outlook no quede en segundo plano antes de la siguiente prueba.
