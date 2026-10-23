# MCExcelBot

A Python automation project that generates structured Excel workbooks and delivers them through a Telegram bot.

The project combines conversational input, validation, reusable spreadsheet generators, formulas, QR identifiers and automated tests. It also includes an experimental image-to-data digitization path with a provider boundary for AI-assisted extraction.

## What it demonstrates

- Telegram bot workflows with `python-telegram-bot`.
- Programmatic Excel generation with `openpyxl`.
- Reusable template registry and generator abstractions.
- Monthly and annual workbook generation.
- Formula generation, totals, accumulated values and input validation.
- Unique document IDs and QR codes embedded into generated workbooks.
- Separation between bot handlers, domain models, services and spreadsheet templates.
- Environment-based configuration for external credentials.
- Automated tests for core generation and bot behavior.

## Architecture

```text
Telegram user
     |
     v
Bot handlers / conversation state
     |
     +------> request validation
     |
     v
Template registry
     |
     +--> Sales / operational workbook
     +--> Productivity workbook
     +--> KPI workbook
     `--> Annual consolidated workbook
              |
              v
       openpyxl generation
        formulas + QR IDs
              |
              v
            .xlsx
              |
              v
      Telegram document reply
```

## Available generators

The repository currently contains four spreadsheet generator families:

- monthly sales and operational tracking;
- productivity and controllable metrics;
- weekly/monthly KPI tracking;
- annual consolidated sales.

The generators share common request/response models and are registered behind a template registry so new workbook types can be added without coupling them directly to Telegram handlers.

## Telegram interface

The bot supports:

- `/start`
- `/crear`
- `/ayuda`
- `/cancelar`

The creation flow uses buttons and validated conversational input to select a template, period and template-specific values before generating the workbook and returning the `.xlsx` file.

## QR document identity

Generated workbooks receive a unique template identifier. A QR payload records the template type, identifier and period, providing a machine-readable link between the physical/printed spreadsheet and its generated identity.

## Optional digitization path

The codebase also contains an AI-assisted digitization workflow intended to extract structured information from spreadsheet images. External AI credentials are supplied through environment variables rather than embedded in source code.

This capability should be treated separately from the deterministic workbook-generation path: generation is the core project feature, while image digitization depends on the configured external provider and input quality.

## Local setup

Requirements:

- Python 3.11+
- a Telegram bot token for bot execution

```bash
python -m venv .venv
pip install -e .
```

Copy the example environment configuration and provide your own credentials:

```bash
cp .env.example .env
```

Then run:

```bash
python scripts/run_bot.py
```

The bot uses long polling for local execution.

## Generate workbooks directly

The spreadsheet engine can also be exercised without Telegram:

```bash
python scripts/generate_sales.py --year 2026 --month 9 --output output
python scripts/generate_productivity.py --year 2026 --month 6 --output output
python scripts/generate_ip.py --year 2026 --month 7 --target 6.5 --output output
python scripts/generate_annual.py --year 2026 --output output
```

## Tests

```bash
python -m unittest discover -s tests -v
```

The repository includes automated coverage for spreadsheet generation, validation and bot-oriented behavior. Test counts are intentionally not hard-coded here because they can change as the project evolves.

## Repository structure

```text
src/bot_excel/
  bot/          # Telegram application, handlers and conversation state
  core/         # domain models and identifiers
  services/     # generation support and digitization services
  templates/    # reusable spreadsheet generators and registry
  branding.py   # QR helpers and optional branding extension point
scripts/        # direct generator and bot entry points
tests/          # automated tests
```

## Portfolio sanitization

This portfolio branch intentionally excludes third-party brand logos from the runtime and keeps branding as an optional extension point. Deployments should provide only assets they own or are authorized to use.

## Scope

MCExcelBot is a technical automation project, not an official product of any third-party brand. The repository demonstrates Python automation, Telegram integration and spreadsheet engineering. External service credentials and organization-specific production data are not required to inspect the source code.
