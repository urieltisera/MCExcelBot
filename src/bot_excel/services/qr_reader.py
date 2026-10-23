from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import cv2
import numpy as np


class QrReaderError(ValueError):
    """Error seguro y presentable al validar o leer un QR."""


class QrNotDetectedError(QrReaderError):
    pass


class InvalidQrPayloadError(QrReaderError):
    pass


class UnsupportedQrTypeError(QrReaderError):
    pass


@dataclass(frozen=True)
class ParsedQr:
    type: str
    code: str
    period: str
    year: int
    month: int
    payload: str


_PAYLOAD_RE = re.compile(
    r"^MCEXCEL\|TIPO=(?P<type>[A-Z_]+)"
    r"\|CODIGO=(?P<code>[A-Z]{2}-\d{4}-\d{2}-[A-Z0-9]{4,12})"
    r"\|PERIODO=(?P<year>\d{4})-(?P<month>\d{2})$"
)


def parse_sales_gc_qr(payload: str) -> ParsedQr:
    match = _PAYLOAD_RE.fullmatch(payload.strip())
    if match is None:
        raise InvalidQrPayloadError("El QR no tiene el formato esperado de MCEXCEL.")

    qr_type = match.group("type")
    if qr_type != "VENTAS_GC":
        raise UnsupportedQrTypeError(
            "Por ahora solo puedo digitalizar planillas de Ventas y GC."
        )

    year = int(match.group("year"))
    month = int(match.group("month"))
    if not 2000 <= year <= 2100 or not 1 <= month <= 12:
        raise InvalidQrPayloadError("El período del QR no es válido.")

    code = match.group("code")
    expected_prefix = f"VT-{year}-{month:02d}-"
    if not code.startswith(expected_prefix):
        raise InvalidQrPayloadError(
            "El código y el período del QR no coinciden."
        )

    return ParsedQr(
        type=qr_type,
        code=code,
        period=f"{year}-{month:02d}",
        year=year,
        month=month,
        payload=payload.strip(),
    )


def decode_sales_gc_qr(image_path: Path) -> ParsedQr:
    path = Path(image_path)
    image_bytes = np.fromfile(path, dtype=np.uint8)
    image = cv2.imdecode(image_bytes, cv2.IMREAD_COLOR)
    if image is None:
        raise QrNotDetectedError("La imagen no se pudo abrir.")

    detector = cv2.QRCodeDetector()
    candidates = [
        image,
        cv2.resize(image, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC),
    ]
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    candidates.extend([
        gray,
        cv2.resize(gray, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC),
    ])

    for candidate in candidates:
        payload, points, _ = detector.detectAndDecode(candidate)
        if points is not None and payload:
            return parse_sales_gc_qr(payload)

    raise QrNotDetectedError(
        "No pude identificar el QR de la planilla."
    )
