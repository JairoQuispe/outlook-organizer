# Outlook Organizer CLI

CLI interactiva para Windows que permite importar archivos PST a Outlook.

## Requisitos

- [Zig 0.15.1](https://ziglang.org/download/)
- PowerShell 7+ (`pwsh`) instalado en el PATH
- Microsoft Outlook instalado

## Estructura

```
apps/cli/
├── build.zig           # Build script
├── src/
│   ├── main.zig        # Punto de entrada, menú principal
│   ├── ui.zig          # Helpers de UI: banner, input, progress bar
│   ├── ps_runner.zig   # Ejecutor de scripts PowerShell embebidos
│   ├── file_browser.zig    # Explorador de archivos PST
│   ├── store_selector.zig  # Selector de buzones de Outlook
│   └── wizards/
│       └── import_pst.zig  # Wizard de importación PST
└── zig-out/bin/
    └── outlook-organizer.exe
```

## Build

### Debug (desarrollo)

```powershell
cd apps/cli
# Copiar scripts PS al directorio src/ (necesario para embed)
Copy-Item ..\..\core\scripts-pwsh\*.ps1 src\
# Compilar
zig build
# Ejecutar
.\zig-out\bin\outlook-organizer.exe
```

### Release (portable)

```powershell
cd apps/cli
Copy-Item ..\..\core\scripts-pwsh\*.ps1 src\
zig build -Doptimize=ReleaseSmall
# El ejecutable portable está en zig-out\bin\outlook-organizer.exe (~380 KB)
```

## Funcionalidades

### Importar PST
1. Explorador de archivos para seleccionar el PST de origen
2. Listado de buzones de Outlook disponibles
3. Selección de acción: Copiar o Mover
4. Opción de saltar duplicados (con revisión profunda opcional)
5. Filtros por año y meses
6. Selección de carpetas específicas a importar
7. Throttling adaptativo para Exchange Online
8. Resumen de configuración antes de ejecutar
9. Resultados detallados al finalizar

### Scan PST
Placeholder — próximamente.

## Notas

- Los scripts PowerShell (`outlook-import-pst.ps1`, `outlook-list-stores.ps1`, `outlook-scan-pst.ps1`) se embeben en el ejecutable en tiempo de compilación.
- El ejecutable es completamente portable, no necesita archivos adicionales.
- Al ejecutar, los scripts se extraen a la carpeta `%TEMP%` y se eliminan al terminar.
