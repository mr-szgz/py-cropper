# Py-Cropper

Py-Cropper is a Typer-powered command-line utility that trims uniform white borders from common bitmap formats using Pillow. It can process single files, batches and (on Windows) integrates with File Explorer via a right-click verb.

## Requirements
- Python 3.8+
- [Pillow](https://python-pillow.org) and [Typer](https://typer.tiangolo.com) (installed automatically when using `pip install`)
- Windows is only required for the optional context-menu verb

## Installation
```powershell
# Install from PyPI (preferred)
pip install py-cropper

# ...or install from a local clone
git clone https://github.com/mr-szgz/py-cropper.git
cd py-cropper
pip install .
```
The installer registers a `cropper` console entry point. On Windows it will also attempt to add the File Explorer verb automatically when `pip install` is executed.

## Usage
```powershell
# Crop a single file (writes <name>-cropped.<ext> by default)
cropper example.png

# Force the output path
cropper example.png --output clean.png

# Process several files and log errors individually
cropper image1.jpg image2.jpg

# Process every supported image located next to <seed>.png
cropper seed.png --include-siblings

# Keep more of the light border by increasing tolerance (0-255)
cropper scan.png --tolerance 230
```
CLI help is always available with `cropper --help` or `cropper --help`.

## Windows File Explorer Shell Extension

### Building Extension
```powershell
dotnet build
```

### Registering in Context Menu

Run the helper from an elevated PowerShell prompt inside `dotnet/` (ensure `cropper.exe` is already on your `PATH`):

```powershell
cd dotnet
.\shell.ps1 -Action Register -Configuration Release
```

This builds the COM host and wires up the Explorer verb. To remove it later run:

```powershell
.\shell.ps1 -Action Unregister
```
