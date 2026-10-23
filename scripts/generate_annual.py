from __future__ import annotations

import argparse
from pathlib import Path

from bot_excel.core.models import AnnualSalesRequest
from bot_excel.templates.ventas_anuales.generator import (
    AnnualSalesTemplateGenerator,
)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Genera la planilla de Ventas consolidadas anuales."
    )
    parser.add_argument("--year", type=int, required=True)
    parser.add_argument("--output", type=Path, default=Path("output"))
    args = parser.parse_args()

    result = AnnualSalesTemplateGenerator().generate(
        AnnualSalesRequest(
            year=args.year,
            output_dir=args.output,
        )
    )
    print(result.workbook_path)


if __name__ == "__main__":
    main()
