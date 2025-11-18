from __future__ import annotations

import logging
import sys
from pathlib import Path

if sys.platform == "win32":
    import winreg
else:  # pragma: no cover - Windows-specific helper
    winreg = None  # type: ignore


_DISPLAY_NAME = "Crop with Py-Cropper"
_BASE_KEY = r"Software\Classes\*\shell\PyCropper"
_COMMAND_KEY = fr"{_BASE_KEY}\command"


def register_context_menu(cropper_executable: Path) -> None:
    """Create/update the Explorer context-menu verb for Py-Cropper."""
    if sys.platform != "win32" or winreg is None:
        logging.info("Skipping context-menu registration; not running on Windows.")
        return

    resolved = cropper_executable.resolve()
    if not resolved.exists():
        raise FileNotFoundError(f"cropper executable not found at {resolved}")

    logging.info("Registering Py-Cropper context menu at %s", resolved)
    command = f'"{resolved}" crop "%1"'

    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _BASE_KEY) as key:
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, _DISPLAY_NAME)

    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, _COMMAND_KEY) as cmd_key:
        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, command)


def unregister_context_menu() -> None:
    """Remove the Explorer verb for Py-Cropper, if present."""
    if sys.platform != "win32" or winreg is None:
        return

    for sub_key in (_COMMAND_KEY, _BASE_KEY):
        try:
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, sub_key)
        except FileNotFoundError:
            continue
