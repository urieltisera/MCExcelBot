from __future__ import annotations

from pathlib import Path

import qrcode
from openpyxl.drawing.image import Image
from openpyxl.worksheet.worksheet import Worksheet


def add_brand_logo(
    sheet: Worksheet,
    anchor: str = "A1",
    width: int = 30,
    height: int = 30,
) -> None:
    """Reserved extension point for caller-owned branding.

    The portfolio version intentionally ships without third-party logos or
    trademark assets. Keeping this function as a no-op preserves the template
    interface while allowing a deployment to provide its own branding layer.
    """
    _ = (sheet, anchor, width, height)


def build_qr_payload(
    template_type: str,
    template_id: str,
    year: int,
    month: int,
) -> str:
    return (
        f"MCEXCEL|TIPO={template_type}|CODIGO={template_id}|"
        f"PERIODO={year}-{month:02d}"
    )


def build_annual_qr_payload(
    template_type: str,
    template_id: str,
    year: int,
) -> str:
    return f"MCEXCEL|TIPO={template_type}|CODIGO={template_id}|PERIODO={year}"


def create_qr_png(payload: str, output_path: Path) -> Path:
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(payload)
    qr.make(fit=True)
    image = qr.make_image(fill_color="black", back_color="white")
    image.save(output_path)
    return output_path


def add_qr_image(
    sheet: Worksheet,
    qr_path: Path,
    anchor: str = "B1",
    width: int = 40,
    height: int = 40,
) -> None:
    image = Image(str(qr_path))
    image.width = width
    image.height = height
    sheet.add_image(image, anchor)


def add_qr_panel(
    sheet: Worksheet,
    qr_path: Path,
    image_anchor: str,
    width: int = 48,
    height: int = 48,
) -> None:
    """Embed a compact QR image without adding visible worksheet text."""
    add_qr_image(
        sheet,
        qr_path,
        anchor=image_anchor,
        width=width,
        height=height,
    )
