from __future__ import annotations

import re
from pathlib import Path

from openpyxl import load_workbook


class GeneratedFileRegistryError(LookupError):
    pass


class GeneratedFileNotFoundError(GeneratedFileRegistryError):
    pass


class MultipleGeneratedFilesError(GeneratedFileRegistryError):
    pass


_SALES_CODE_RE = re.compile(r"^VT-\d{4}-\d{2}-[A-Z0-9]{4,12}$")


class GeneratedFileRegistry:
    def __init__(self, output_dir: Path) -> None:
        self.output_dir = Path(output_dir).resolve()

    def find_sales_gc(self, template_code: str) -> Path:
        if _SALES_CODE_RE.fullmatch(template_code) is None:
            raise GeneratedFileNotFoundError("El código de planilla no es válido.")

        suffix = f"_{template_code}.xlsx"
        matches = [
            path.resolve()
            for path in self.output_dir.glob("Ventas_GC_*.xlsx")
            if path.name.endswith(suffix)
        ]
        matches = [path for path in matches if self._is_inside_output(path)]

        if not matches:
            raise GeneratedFileNotFoundError(
                "No encontré el Excel original correspondiente al QR."
            )
        if len(matches) > 1:
            raise MultipleGeneratedFilesError(
                "Hay más de un Excel original con el mismo código."
            )
        match = matches[0]
        try:
            workbook = load_workbook(match, read_only=True, data_only=False)
            config = workbook["Configuración"]
            stored_code = config["B6"].value
            payload = config["B7"].value
            workbook.close()
        except Exception as exc:
            raise GeneratedFileNotFoundError(
                "El Excel original no es una planilla válida de Ventas y GC."
            ) from exc
        if stored_code != template_code or not str(payload).startswith(
            "MCEXCEL|TIPO=VENTAS_GC|"
        ):
            raise GeneratedFileNotFoundError(
                "El Excel encontrado no coincide con el tipo y código del QR."
            )
        code_parts = template_code.split("-")
        expected_payload = (
            "MCEXCEL|TIPO=VENTAS_GC|"
            f"CODIGO={template_code}|"
            f"PERIODO={code_parts[1]}-{code_parts[2]}"
        )
        if payload != expected_payload:
            raise GeneratedFileNotFoundError(
                "El período interno del Excel no coincide con el QR."
            )
        return match

    def _is_inside_output(self, path: Path) -> bool:
        try:
            path.relative_to(self.output_dir)
        except ValueError:
            return False
        return path.is_file()
