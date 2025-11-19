# Cropper Shell Extension

Windows shell extension that exposes `cropper.exe` directly in both the modern (Windows 11) and classic Explorer context menus. The handler forwards the first selected file path to PyCropper, so make sure `cropper.exe` is on your `PATH` (e.g., install the Python package and `mise`/pipx puts it there).

## Build & Register

Run the helper script from an elevated PowerShell prompt inside `dotnet/`:

```powershell
.\register-extension.ps1 -Action Register -Configuration Release
```

This publishes the COM host (`net8.0-windows`, `win-x64`) and calls `regsvr32` for you. After the popup confirmation, right-click any file to see **Crop with PyCropper**.

## Unregister

```powershell
.\register-extension.ps1 -Action Unregister
```

That removes both the modern Explorer command and the legacy `shellex` handler.
