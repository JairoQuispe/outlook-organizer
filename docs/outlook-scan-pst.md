# Guía rápida para `outlook-scan-pst.ps1`

> Script: `scripts-powershell/outlook-scan-pst.ps1`
> Ejecuta un escaneo de carpetas dentro de un PST montado y expone conteos por año, tamaños y métricas resumidas.

## Uso general

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 [-PstPath <ruta>] [-StoreId <id>] [flags]
```

- Provee `-PstPath` para montar un archivo PST o `-StoreId` para apuntar a un store ya montado.
- El script intenta reusar sesiones de Outlook abiertas a menos que actives `-PreserveSession`.
- Genera eventos `scanMeta`, `scanProgress` y una colección plana de carpetas (JSON) que consume la UI.

## Flags disponibles

| Flag | Tipo | Descripción |
|------|------|-------------|
| `-PstPath` | `string` | Ruta directa al PST que se montará para el escaneo. Si falta, debes pasar `-StoreId`.
| `-StoreId` | `string` | Identificador de store existente (extraído con `outlook-list-stores.ps1`).
| `-FilterOnlyYear` | `int` | Limita los ítems contados al año exacto; omite carpetas sin coincidencias para ese año.
| `-IncludeSize` | `switch` | Calcula `sizeBytes`/`sizeHuman` usando la columna `Size` en la tabla MAPI (cuando está disponible) y lo añade por carpeta.
| `-Summary` | `switch` | Genera al final un JSON único `type: "summary"` con inputs, fuentes, totales, year breakdown y tops (no repite el dump plano).
| `-ExportResult [json|text]` | `string` | (Opcional) Escribe el payload final en `scan-result-<timestamp>.json` (default) o `.txt` si pones `text`. El modo texto formatea el resumen con saltos de línea.
| `-ExportFolders` | `switch` | Exporta solo la lista de carpetas filtradas a un archivo JSON. El nombre incluye los flags usados (año, store, tamaño, fecha/hora). Ideal para ser consumido por otro script PowerShell.
| `-ExportFoldersPath` | `string` | (Opcional) Ruta personalizada para el archivo generado por `-ExportFolders`. Si se omite, se guarda en el directorio actual.
| `-Json` | `switch` | Salida compacta en JSON para logs y automatización.
| `-Headless` | `switch` | Igual que `-Json` pero más explícito para scripts UILess.
| `-PreserveSession` | `switch` | No cierra ni desmonta Outlook al terminar, ideal cuando ya tienes una instancia abierta.

## Flujo típico y ejemplos

1. Escanear un PST en modo normal (JSON + carpetas):

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -PstPath "C:\PSTs\archivo.pst" -Json
```

Salida parcial:
```json
{"type":"scanMeta","pstPath":"C:\\PSTs\\archivo.pst","totalFolders":124,...}
{"type":"folders","count":124}
{"type":"folder","path":"Inbox","itemCount":1024,...}
...``` 

2. Incluir tamaños y filtrar por año: 

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -StoreId "00000000C..." -FilterOnlyYear 2023 -IncludeSize -Json
```

Cada carpeta emitida ahora contiene `sizeBytes` y `sizeHuman` que reflejan los ítems considerados.

3. Obtener resumen y exportar a disco:

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -PstPath "C:\PSTs\archivo.pst" -Summary -ExportResult text
```

- Se seguirá imprimiendo el mismo `scanMeta` + `scanProgress` hasta el final.
- Al terminar, se escribe `scan-result-<timestamp>.txt` con un resumen legible (inputs, totales, top folders, etc.).

4. Exportar el resumen directo en JSON (ideal para otra app):

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -StoreId "00000000C..." -Summary -ExportResult json
```

El archivo `scan-result-<timestamp>.json` contiene:
```json
{
  "type": "summary",
  "inputs": {...},
  "source": {...},
  "scan": {...},
  "totals": {...},
  "yearBreakdown": [...],
  "topFoldersByItems": [...],
  "topFoldersBySize": [...],
  "folders": [ ... lista completa ... ]
}
```

5. Exportar solo la lista de carpetas filtradas:

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -PstPath "C:\PSTs\archivo.pst" -FilterOnlyYear 2023 -IncludeSize -ExportFolders
```

Genera un archivo como `scan-year-2023-pst-archivo-with-size-20250512-183000.json` con:
```json
{
  "type": "folderExport",
  "generatedAt": "2025-05-12T18:30:00.000Z",
  "pstPath": "C:\\PSTs\\archivo.pst",
  "storeId": null,
  "filterOnlyYear": 2023,
  "includeSize": true,
  "folderCount": 42,
  "folders": [
    { "type": "folder", "path": "Inbox", "itemCount": 150, "sizeBytes": 5242880, "sizeHuman": "5.00 MB", ... },
    ...
  ]
}
```

6. Exportar carpetas a una ruta personalizada:

```powershell
pwsh -ExecutionPolicy Bypass -File scripts-powershell/outlook-scan-pst.ps1 -StoreId "00000000C..." -ExportFolders -ExportFoldersPath "D:\exports\carpetas.json"
```

## Notas adicionales

- Ejecutar sin `-PstPath`/`-StoreId` aborta con error inmediato. 
- `-Summary` recalcula totales y desgloses, pero no elimina los eventos individuales para retrocompatibilidad.
- `-ExportResult` respeta el último payload emitido (summary o listado) y agrega logging del archivo generado.
- `-ExportFolders` genera un JSON independiente con metadata del filtro usado, ideal para encadenar con otro script PS que lo lea con `Get-Content | ConvertFrom-Json`.
- Para recolectar `StoreId`, usa `scripts-powershell/outlook-list-stores.ps1 -Json`.
