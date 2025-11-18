from __future__ import annotations

import sys
from pathlib import Path

from setuptools import setup
from setuptools.command.install import install


class InstallWithRegistryVerb(install):
    def run(self) -> None:
        super().run()
        self._register_context_menu()

    def _register_context_menu(self) -> None:
        if sys.platform != "win32":
            return

        scripts_dir = getattr(self, "install_scripts", None)
        if not scripts_dir:
            return

        candidate = Path(scripts_dir) / "cropper.exe"
        if not candidate.exists():
            alt = Path(scripts_dir) / "cropper"
            candidate = alt if alt.exists() else candidate

        if not candidate.exists():
            return

        from context_menu import register_context_menu

        try:
            register_context_menu(candidate)
        except Exception:
            pass


setup(cmdclass={"install": InstallWithRegistryVerb})
