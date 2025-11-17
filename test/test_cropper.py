from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from PIL import Image, ImageDraw
from typer.testing import CliRunner

from cropper import app, auto_crop

ROOT = Path(__file__).resolve()
SAMPLE_DIR = ROOT / "sample"
runner = CliRunner()

SAMPLE_CASES = (
    ("white-border-portrait.jpg", (1041, 1500)),
    ("white-flower-model-left+right.png", (340, 191)),
    ("white-mannequin-left+right.jpg", (1007, 1515)),
)


def _copy_sample(tmp_path: Path, filename: str) -> Path:
    """Copy one of the sample assets to the pytest tmp_path."""
    source = SAMPLE_DIR / filename
    dest = tmp_path / filename
    shutil.copyfile(source, dest)
    return dest


@pytest.mark.parametrize(("filename", "expected_size"), SAMPLE_CASES)
def test_auto_crop_trims_white_borders(tmp_path, filename, expected_size):
    source = _copy_sample(tmp_path, filename)
    destination = tmp_path / f"{source.stem}-cropped{source.suffix}"

    written = auto_crop(source, destination)

    assert written == destination
    assert destination.exists()
    with Image.open(destination) as cropped:
        assert cropped.size == expected_size


def test_auto_crop_uses_default_destination(tmp_path):
    source = tmp_path / "bordered.png"
    image = Image.new("RGB", (20, 20), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((5, 5, 14, 14), fill="black")
    image.save(source)

    written = auto_crop(source)

    expected = source.with_name("bordered-cropped.png")
    assert written == expected
    assert expected.exists()
    with Image.open(expected) as cropped:
        assert cropped.size == (10, 10)


def test_auto_crop_rejects_completely_white_images(tmp_path):
    source = tmp_path / "all-white.png"
    Image.new("RGB", (8, 8), "white").save(source)

    with pytest.raises(ValueError):
        auto_crop(source)


def test_cli_crop_command(tmp_path):
    source = _copy_sample(tmp_path, "white-mannequin-left+right.jpg")
    expected = source.with_name(f"{source.stem}-cropped{source.suffix}")

    result = runner.invoke(app, [str(source)])
    assert result.exit_code == 0, result.stdout
    assert "Cropped image written to" in result.stdout
    assert expected.exists()
    with Image.open(expected) as cropped:
        assert cropped.size == (1007, 1515)
