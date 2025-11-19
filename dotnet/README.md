# Cropper Shell Extension

Windows shell extension that exposes `cropper.exe` directly in both the modern (Windows 11) and classic Explorer context menus. The handler forwards the first selected file path to PyCropper, so make sure `cropper.exe` is on your `PATH` (e.g., install the Python package and `mise`/pipx puts it there).

## Build & Register

Run the helper script from an elevated PowerShell prompt inside `dotnet/`:

```powershell
.\shell.ps1 -Action Install -Configuration Release
```

The `Install` action automatically unregisters the existing shell extension before rebuilding/publishing (`net8.0-windows`, `win-x64`) and re-registering via `regsvr32`, so Explorer always gets the latest bits. After the popup confirmation, right-click any file to see **Crop with PyCropper**.

Use `-Force` if you want the pre-unregister step while invoking `-Action Register` or to ignore missing binaries while removing the extension.

## Unregister

```powershell
.\shell.ps1 -Action Unregister
```

That removes both the modern Explorer command and the legacy `shellex` handler.
