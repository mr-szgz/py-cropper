from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Tuple

from importlib import metadata
import typer
from PIL import Image, ImageOps

app = typer.Typer(help="Auto-crop white borders away using Pillow.")

IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".bmp",
    ".gif",
    ".tiff",
    ".webp",
}


def _read_local_version() -> str | None:
    pyproject = Path(__file__).with_name("pyproject.toml")
    if not pyproject.exists():
        return None

    inside_project = False
    for raw_line in pyproject.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        if line == "[project]":
            inside_project = True
            continue
        if inside_project and line.startswith("["):
            break

        if inside_project:
            match = re.match(r'version\s*=\s*"([^"]+)"', line)
            if match:
                return match.group(1)
    return None


def _resolve_version() -> str:
    for candidate in ("pycropper", "cropper"):
        try:
            return metadata.version(candidate)
        except metadata.PackageNotFoundError:
            continue
    local = _read_local_version()
    return local or "unknown"


APP_VERSION = _resolve_version()


def _version_callback(value: bool) -> None:
    if not value:
        return

    typer.echo(f"pycropper {APP_VERSION}")
    raise typer.Exit()


def auto_crop(
    source: Path,
    destination: Path | None = None,
    tolerance: int = 5,
) -> Path | None:
    """
    uninstall white borders from ``source`` using Pillow's optimized helpers.

    Args:
        source: Image file to crop.
        destination: Where to store the cropped image. Defaults to ``<name>-cropped``.
        tolerance: How close a pixel can be to pure white (0-255 scale) before
            it is treated as background.

    Returns:
        Path to the written file, or ``None`` if no cropping was performed.
    """

    img = Image.open(source).convert("RGB")
    dest = destination or source.with_name(f"{source.stem}-cropped{source.suffix}")

    _, bbox = _threshold_mask(img, tolerance)
    if bbox is None:
        raise ValueError("Unable to determine crop box - is the image entirely uniform?")

    if bbox == (0, 0, img.width, img.height):
        return None

    cropped = img.crop(bbox)
    cropped.save(dest)
    return dest


def _threshold_mask(
    img: Image.Image,
    tolerance: int,
) -> Tuple[Image.Image, Tuple[int, int, int, int] | None]:
    gray = ImageOps.grayscale(img)
    tolerance = max(0, min(255, tolerance))
    limit = max(0, min(255, 255 - tolerance))

    def _threshold(pixel: int, thresh: int = limit) -> int:
        return 255 if int(pixel) < thresh else 0

    mask = gray.point(_threshold)
    return mask, mask.getbbox()


def _list_image_files(directory: Path) -> list[Path]:
    return sorted(
        file
        for file in directory.iterdir()
        if file.is_file() and file.suffix.lower() in IMAGE_EXTENSIONS
    )


def _run_crop_cli(
    sources: list[Path] = typer.Argument(
        ...,
        exists=True,
        dir_okay=False,
        readable=True,
        resolve_path=True,
        help="Input image file(s).",
    ),
    include_siblings: bool = typer.Option(
        False,
        "--include-siblings",
        "--siblings",
        "-s",
        is_flag=True,
        help="Also process every supported image in the source folder.",
    ),
    output: Path | None = typer.Option(
        None,
        "--output",
        "-o",
        dir_okay=False,
        writable=True,
        resolve_path=True,
        help="Destination file. Defaults to <name>-cropped.<ext>.",
    ),
    tolerance: int = typer.Option(
        200,
        "--tolerance",
        "-t",
        min=0,
        max=255,
        help="Background tolerance (0-255). Higher tolerances keep more border.",
    ),
) -> None:
    if include_siblings and output is not None:
        raise typer.BadParameter("--output cannot be used together with --include-siblings.")
    if len(sources) > 1 and output is not None:
        raise typer.BadParameter("--output can only be used when a single source is provided.")

    if include_siblings:
        if len(sources) != 1:
            raise typer.BadParameter("--include-siblings requires exactly one source.")
        source = sources[0]
        images = _list_image_files(source.parent)
        if not images:
            typer.echo(f"No supported images found in {source.parent}")
            raise typer.Exit(code=1)
        for path in images:
            try:
                written = auto_crop(path, None, tolerance)
            except Exception as exc:  # pragma: no cover - runtime feedback
                typer.echo(f"[ERROR] {path.name}: {exc}")
            else:
                if written is None:
                    typer.echo(f"Skipped {path.name} (no border detected)")
                else:
                    typer.echo(f"Wrote {written}")
        return

    if len(sources) == 1:
        written = auto_crop(sources[0], output, tolerance)
        if written is None:
            typer.echo("No borders detected; nothing written.")
        else:
            typer.echo(f"Cropped image written to {written}")
        return

    had_error = False
    for path in sources:
        try:
            written = auto_crop(path, None, tolerance)
        except Exception as exc:  # pragma: no cover - runtime feedback
            typer.echo(f"[ERROR] {path.name}: {exc}")
            had_error = True
            continue

        if written is None:
            typer.echo(f"{path.name}: No borders detected; nothing written.")
        else:
            typer.echo(f"{path.name}: Cropped image written to {written}")

    if had_error:
        raise typer.Exit(code=1)


@app.command("crop")
def crop_command(
    _show_version: bool = typer.Option(
        False,
        "--version",
        callback=_version_callback,
        is_flag=True,
        is_eager=True,
        help="Show version and exit.",
    ),
    sources: list[Path] = typer.Argument(
        ...,
        exists=True,
        dir_okay=False,
        readable=True,
        resolve_path=True,
        help="Input image file(s).",
    ),
    include_siblings: bool = typer.Option(
        False,
        "--include-siblings",
        "--siblings",
        "-s",
        is_flag=True,
        help="Also process every supported image in the source folder.",
    ),
    output: Path | None = typer.Option(
        None,
        "--output",
        "-o",
        dir_okay=False,
        writable=True,
        resolve_path=True,
        help="Destination file. Defaults to <name>-cropped.<ext>.",
    ),
    tolerance: int = typer.Option(
        200,
        "--tolerance",
        "-t",
        min=0,
        max=255,
        help="Background tolerance (0-255). Higher tolerances keep more border.",
    ),
) -> None:
    _run_crop_cli(sources, include_siblings, output, tolerance)


def main(argv: list[str] | None = None) -> None:
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] == "crop":
        app()
        return

    typer.run(crop_command)


if __name__ == "__main__":
    main()
