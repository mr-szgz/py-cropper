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

## Uninstall
```powershell
pip uninstall py-cropper
```
Uninstalling removes the CLI entry point. If you previously added the Explorer verb, remove it manually (see below) before uninstalling or reinstall directly from source to re-register it.

## Context Menu (Windows)
Py-Cropper ships with helpers to create a `Crop with Py-Cropper` entry when you right-click an image in File Explorer.

```powershell
# Install the verb (requires existing cropper.exe/scripts entry)
cropper context-menu install

# Remove the verb
cropper context-menu remove
```
Optional switches:
- `--executable PATH` – run a custom `cropper` executable/script when the verb fires (defaults to the current `cropper` command).

If the verb ever breaks because the executable moved, re-run `cropper context-menu install --executable <new-path>` or uninstall/reinstall the package to refresh the registry keys.

> [!WARNING]
> This package was primarily developed with AI assistants: GPT-5, GPT-5-Codex, GPT-5.1, and GPT-5.1-Codex. Reasoning mode stays on Medium by default—High reasoning has consistently produced worse outcomes for this workflow.
