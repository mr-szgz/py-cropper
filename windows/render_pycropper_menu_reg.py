import os
from jinja2 import Template
from shutil import which
from pathlib import Path

base_dir = Path(__file__).resolve().parent
template_path = base_dir / "PyCropperMenu.reg.jinja2"

computer_name = os.environ.get("COMPUTERNAME")
if not computer_name:
    raise EnvironmentError("COMPUTERNAME environment variable is not set.")
output_path = base_dir / f"PyCropperMenu.{computer_name}.reg"

def find_cropper_executable() -> str:
    """Prefer the real executable, but fall back to any runnable shim."""
    for candidate in ("cropper.exe", "cropper"):
        located = which(candidate)
        if not located:
            continue
        resolved = Path(located).resolve()
        if resolved.is_file():
            return os.path.normpath(str(resolved))
    raise FileNotFoundError("cropper executable not found in PATH.")


cropper_path = find_cropper_executable()
tolerances = [50, 100, 150, 200, 250]
# Registry files need backslashes escaped
reg_cropper_path = cropper_path.replace("\\", "\\\\")


with open(template_path, "r", encoding="utf-8") as f:
    tpl = Template(f.read())

rendered = tpl.render(
    cropper_path=reg_cropper_path,
    tolerances=tolerances,
)

with open(output_path, "w", encoding="utf-8") as f:
    f.write(rendered)
