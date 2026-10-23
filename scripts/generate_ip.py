from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from bot_excel.core.models import MonthlyTemplateRequest
from bot_excel.templates.registry import TemplateRegistry


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generar planilla de IP semanal y mensual"
    )
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--target", type=float, default=6.5)
    parser.add_argument("--output", type=Path, default=Path("output"))
    parser.add_argument("--template-id", default=None)
    args = parser.parse_args()

    generator = TemplateRegistry().get("ip_semanal_mensual")
    result = generator.generate(MonthlyTemplateRequest(
        year=args.year,
        month=args.month,
        output_dir=args.output,
        ip_target=args.target,
        template_id=args.template_id,
    ))
    print(result.workbook_path)


if __name__ == "__main__":
    main()
