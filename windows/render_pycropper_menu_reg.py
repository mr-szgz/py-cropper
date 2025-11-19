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

p = which("cropper")
cropper_path = os.path.normpath(str(Path(p).resolve())) if p else None
if cropper_path is None:
    raise FileNotFoundError("cropper executable not found in PATH.")

tolerances = [50, 100, 150, 200, 250]

with open(template_path, "r", encoding="utf-8") as f:
    tpl = Template(f.read())

rendered = tpl.render(
    cropper_path=cropper_path,
    tolerances=tolerances,
)

with open(output_path, "w", encoding="utf-8") as f:
    f.write(rendered)
