# Scripts de PowerShell para Outlook Organizer (`core/scripts-pwsh`)

Esta carpeta contiene el motor principal de automatización de Microsoft Outlook, desarrollado en PowerShell (`.ps1`). Estos scripts interactúan directamente con la API COM de Outlook (MAPI) para realizar consultas, escaneos e importaciones masivas de datos de manera eficiente, segura y altamente configurable.

Todos los scripts están diseñados para ser autónomos o para integrarse de forma fluida con una interfaz de línea de comandos (CLI) escrita en Zig, utilizando formatos de salida estándar como texto plano o JSON estructurado para el reporte de progreso en tiempo real.

---

## Índice de Contenidos
- [Requisitos y Configuración](#requisitos-y-configuración)
- [1. outlook-list-stores.ps1](#1-outlook-list-storesps1) (Listado de Buzones/Stores)
- [2. outlook-scan-pst.ps1](#2-outlook-scan-pstps1) (Escaneo de Archivos PST)
- [3. outlook-import-pst.ps1](#3-outlook-import-pstps1) (Importación/Migración con Throttling)
- [Políticas de Reintento y Resiliencia](#políticas-de-reintento-y-resiliencia)

---

## Requisitos y Configuración

1. **Sistema Operativo**: Windows con Microsoft Outlook instalado y registrado (soporte completo de COM API).
2. **Permisos de Ejecución**: Los scripts configuran de manera automática la política de ejecución local como `Bypass` para el proceso actual:
   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
   ```
3. **Integración**: Los scripts utilizan codificación `UTF-8` al activarse el modo `--Json` o `--Headless`, permitiendo una comunicación perfecta a través de tuberías (pipes) de entrada/salida estándar sin corromper caracteres especiales (tildes, eñes).

---

## 1. outlook-list-stores.ps1

### Descripción
Este script se encarga de enumerar e inspeccionar de forma segura todos los almacenes de datos (*Stores*) conectados al perfil activo de Outlook. Permite identificar buzones de Exchange Online (Office 365), archivos de carpetas locales fuera de línea (`.ost`) y archivos de datos personales (`.pst`).

### Características Clave
- **Clasificación Automática**: Determina con precisión el tipo de buzón/archivo analizando las propiedades internas de Exchange y la extensión del archivo local.
- **Detección de Sesión Activa**: Detecta si Outlook ya se encuentra ejecutándose para evitar conflictos de exclusión de perfiles de correo MAPI.
- **Múltiples Modos de Salida**: Puede imprimir líneas sencillas en la consola para lectura humana o generar payloads JSON compactos listos para análisis automatizados.

### Parámetros
- `-Json` *(Switch)*: Devuelve el resultado como un string JSON estructurado en la salida estándar en lugar de texto legible para humanos.
- `-PreserveSession` *(Switch)*: Indica que se debe reutilizar la sesión existente de Outlook y evita iniciar/cerrar el proceso de Outlook en segundo plano si ya estaba abierto.
- `-StoreType` *(ValidateSet: 'ExchangeOnline', 'OST', 'PST')*: Filtra la lista de almacenes de datos devueltos para mostrar únicamente el tipo seleccionado.
- `-ExportResult` *(ValidateSet: 'None', 'Json', 'Text'; Default: 'None')*: Si se especifica, guarda los resultados enumerados en un archivo (`list-store-yyyyMMdd-HHmmss.json` o `.txt`) dentro del directorio de trabajo actual.

### Ejemplo de Uso
```powershell
# Listar todos los archivos PST conectados en formato JSON por salida estándar
.\outlook-list-stores.ps1 -StoreType PST -Json
```

---

## 2. outlook-scan-pst.ps1

### Descripción
Realiza un escaneo recursivo exhaustivo de un archivo `.pst` (o de un *Store* existente mediante su identificador único). Este script mapea la estructura de carpetas, cuantifica los correos y elementos contenidos y genera un desglose de volumen por año. Es la pieza fundamental utilizada en el asistente de importación para previsualizar los contenidos del archivo y planificar la migración.

### Características Clave
- **Estructura Plana y Jerárquica**: Genera listas detalladas que incluyen la ruta absoluta de las carpetas internas de Outlook.
- **Desglose Temporal**: Clasifica los elementos por año de recepción/creación, permitiendo filtrados temporales precisos.
- **Reportes de Progreso en Tiempo Real**: Envía eventos estructurados en JSON (`scanProgress` con fases, porcentajes y velocidad) para que la CLI que lo llama pueda renderizar barras de progreso fluidas.
- **Cálculo de Tamaños**: Opcionalmente recopila el tamaño en bytes de cada elemento y carpeta.
- **Estadísticas de Uso**: Identifica carpetas con mayor concentración de correos y los mayores consumidores de almacenamiento.

### Parámetros
- `-PstPath` *(String)*: Ruta absoluta del archivo `.pst` a escanear en el sistema de archivos.
- `-StoreId` *(String)*: Identificador de almacén alternativo (MAPI StoreID) si se prefiere escanear un buzón ya montado en el perfil.
- `-FilterOnlyYear` *(Int)*: Filtra el conteo de elementos para restringir el análisis a un único año específico (rango válido: 1900 a 2100).
- `-IncludeSize` *(Switch)*: Habilita el cálculo del tamaño acumulado de cada carpeta en bytes (puede incrementar ligeramente el tiempo de ejecución en PSTs de gran volumen).
- `-Summary` *(Switch)*: En lugar de listar cada carpeta individualmente, genera un payload estadístico global estructurado con resúmenes de distribución anual, carpetas top en elementos y carpetas top en tamaño.
- `-ExportResult` *(ValidateSet: 'json', 'text'; Default: 'json')*: Formato del reporte que será guardado físicamente en disco.
- `-ExportFolders` *(Switch)*: Exporta la estructura completa de carpetas mapeadas a un archivo JSON físico en disco.
- `-ExportFoldersPath` *(String)*: Ruta personalizada donde guardar el archivo JSON de carpetas exportadas.
- `-Json` / `-Headless` *(Switches)*: Fuerzan la salida estándar a emitir únicamente payloads JSON estructurados (incluidos logs y progreso).
- `-PreserveSession` *(Switch)*: Evita interferir con la instancia principal de Outlook en ejecución.

### Ejemplo de Uso
```powershell
# Obtener un resumen estadístico detallado de un PST en disco con conteo de tamaños
.\outlook-scan-pst.ps1 -PstPath "C:\Data\archivo_2023.pst" -IncludeSize -Summary -Json
```

---

## 3. outlook-import-pst.ps1

### Descripción
Este es el motor de migración e importación masiva del proyecto. Copia o mueve correos y carpetas estructuradas desde un archivo `.pst` de origen hacia un buzón destino de Outlook/Exchange Online. Ha sido optimizado con algoritmos de control de flujo para prevenir bloqueos por saturación de peticiones COM y cuotas del servidor de Exchange.

### Características Clave
- **Algoritmo Token Bucket**: Implementa una limitación estricta de velocidad (*throttling*) configurable mediante elementos por minuto y ráfagas (*burst*), protegiendo la sesión contra bloqueos de MAPI de red.
- **Throttling Adaptativo**: Monitorea de forma continua las respuestas COM; si detecta problemas de red o demoras en la respuesta del servidor Exchange, disminuye automáticamente el ritmo de importación y se recupera gradualmente tras periodos de éxito.
- **Detección Avanzada de Duplicados**: Ofrece una opción de comparación profunda (*Deep Duplicate Check*) basada en metadatos clave (Asunto, Remitente, Hash del cuerpo y Fecha de Recepción) para evitar la duplicidad de correos aun cuando carezcan de IDs de mensaje estables.
- **Bypass de Límites de Tamaño**: Filtra de forma preventiva adjuntos y correos que superen el tamaño máximo admitido por Exchange (`MaxItemSizeBytes`).
- **Resiliencia de Conexión**: En caso de pérdida temporal de conexión de red o caída RPC MAPI, suspende temporalmente el proceso, realiza esperas exponenciales (*exponential backoff*) y reestablece la conexión de forma transparente sin perder el progreso.

### Parámetros Principales
- `-PstPath` *(String)*: Ruta física del archivo `.pst` de origen.
- `-TargetStoreId` *(String)*: ID del buzón destino en el cual se importarán los correos.
- `-Action` *(ValidateSet: 'Copy', 'Move'; Default: 'Copy')*: Acción a realizar sobre los correos analizados.
- `-FilterOnlyYear` *(Int)*: Importa únicamente correos pertenecientes a un determinado año.
- `-FilterOnlyMonths` *(String)*: Filtra la importación por meses específicos mediante una lista separada por comas (admite números del 1 al 12 y nombres de meses en inglés/español, ej. `"enero,febrero,oct"`).
- `-SkipDuplicates` *(Switch)*: Compara elementos para no importar aquellos que ya existan en la carpeta de destino.
- `-DeepDuplicateCheck` *(Switch)*: Habilita el análisis exhaustivo de propiedades de correo para la deduplicación estricta.
- `-ItemsPerMinute` *(Int; Default: 120)*: Tasa de velocidad máxima de transferencia admitida.
- `-BurstSize` *(Int; Default: 20)*: Tamaño de ráfaga máxima que se puede procesar sin esperas de control.
- `-AdaptiveThrottling` *(Switch)*: Activa el ajuste dinámico e inteligente de velocidad basado en latencia y errores.
- `-FolderPlanPath` *(String)*: Ruta a un archivo JSON que contiene la planeación predefinida de la estructura de carpetas a migrar (útil para migraciones estructuradas complejas).
- `-IncludeFolders` / `-IncludeFoldersJson` *(String[] / String)*: Filtros para procesar únicamente subconjuntos específicos de carpetas.
- `-Json` / `-Headless` *(Switches)*: Habilitan la salida de telemetría estructurada de rendimiento y progreso para su interpretación por la CLI Zig.

### Ejemplo de Uso
```powershell
# Importar correos del año 2022 deduplicando profundamente, con tasa máxima de 150 correos/min y control adaptativo
.\outlook-import-pst.ps1 -PstPath "C:\Data\archivo.pst" -TargetStoreId "00000000..." -Action Copy -FilterOnlyYear 2022 -SkipDuplicates -DeepDuplicateCheck -ItemsPerMinute 150 -AdaptiveThrottling -Json
```

---

## Políticas de Reintento y Resiliencia

La comunicación a través del protocolo COM de Outlook en Windows suele ser sensible a la carga de trabajo y a la interacción de red (en el caso de Exchange Online). Por ello, los scripts de este repositorio comparten políticas avanzadas:

1. **Reintentos en Casos de Ocupado (Busy State)**: Cuando la COM API arroja códigos como `0xCC54011D` (Outlook ocupado), los scripts detienen la ejecución, realizan una pausa en segundos (configurada en `$retryDelay`) y reintentan la acción hasta en 3 o 5 ocasiones consecutivas antes de fallar.
2. **Liberación Estricta de Memoria COM**: La API COM es propensa a fugas de memoria y bloqueos de archivos en disco. El motor de scripts implementa de forma agresiva la liberación inmediata de punteros de objetos a través del método de liberación `[System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj)` y disparando llamadas al recolector de basura de .NET (`[System.GC]::Collect()`) para asegurar el cierre inmediato del archivo `.pst` al culminar los escaneos.
3. **Mantenimiento de Sesión**: Los comandos detectan instancias de Outlook huérfanas o preexistentes, decidiendo de manera autónoma si levantar un servicio de fondo (`headless`) o acoplarse al cliente de Outlook que el usuario tiene abierto para asegurar una experiencia silenciosa y eficiente.
