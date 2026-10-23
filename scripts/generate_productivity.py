from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.templates.registry import TemplateRegistry


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generar planilla mensual de Productividad y controlables"
    )
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--output", type=Path, default=Path("output"))
    parser.add_argument("--template-id", default=None)
    args = parser.parse_args()

    generator = TemplateRegistry().get("productividad_controlables")
    result = generator.generate(MonthlyTemplateRequest(
        year=args.year,
        month=args.month,
        output_dir=args.output,
        template_id=args.template_id,
    ))
    print(result.workbook_path)


if __name__ == "__main__":
    main()
