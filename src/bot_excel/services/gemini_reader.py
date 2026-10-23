from __future__ import annotations

from pathlib import Path
from typing import Any, Callable

from pydantic import ValidationError

from bot_excel.core.digitization_models import SalesGcPhotoReading


class GeminiReaderError(RuntimeError):
    pass


class GeminiConfigurationError(GeminiReaderError):
    pass


class GeminiResponseError(GeminiReaderError):
    pass


EXTRACTION_PROMPT = """
Analizá esta fotografía de una planilla mensual de Ventas y GC.

Extraé únicamente los campos manuscritos o completados de las columnas:
- VENTA REAL
- GC REAL

Reglas obligatorias:
- No leas ni recalcules Venta proyectada, GC proyectados, Diferencia,
  Acumulado ni Totales.
- No infieras valores faltantes.
- Una celda vacía debe ser null, nunca cero.
- Un cero escrito explícitamente sí debe ser 0.
- Respetá el número de día impreso en cada fila.
- No inventes valores fuera de cuadro o ilegibles.
- Las confianzas deben representar la legibilidad de cada valor entre 0 y 1.
- Advertí sobre foto borrosa, recortes, reflejos, escritura ambigua,
  filas fuera de cuadro y números posiblemente confundidos.
- Devolvé exclusivamente el esquema estructurado solicitado.
""".strip()


class GeminiSalesGcReader:
    def __init__(
        self,
        *,
        api_key: str | None,
        model: str = "gemini-3.5-flash-lite",
        client_factory: Callable[..., Any] | None = None,
    ) -> None:
        self._api_key = (api_key or "").strip()
        self.model = model
        self._client_factory = client_factory

    def read_image(
        self,
        image_path: Path,
        *,
        mime_type: str,
    ) -> SalesGcPhotoReading:
        if not self._api_key:
            raise GeminiConfigurationError(
                "La digitalización no está configurada en este momento."
            )
        if mime_type not in {"image/jpeg", "image/png", "image/webp"}:
            raise GeminiReaderError("El formato de imagen no está permitido.")

        image_bytes = Path(image_path).read_bytes()
        client = None
        try:
            if self._client_factory is None:
                from google import genai

                client = genai.Client(api_key=self._api_key)
            else:
                client = self._client_factory(api_key=self._api_key)

            from google.genai import types

            response = client.models.generate_content(
                model=self.model,
                contents=[
                    types.Part.from_text(text=EXTRACTION_PROMPT),
                    types.Part.from_bytes(data=image_bytes, mime_type=mime_type),
                ],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    response_schema=SalesGcPhotoReading,
                    seed=1,
                    max_output_tokens=8192,
                ),
            )
            if isinstance(response.parsed, SalesGcPhotoReading):
                return response.parsed
            if isinstance(response.parsed, dict):
                return SalesGcPhotoReading.model_validate(response.parsed)
            if response.text:
                return SalesGcPhotoReading.model_validate_json(response.text)
            raise GeminiResponseError(
                "Gemini no devolvió una lectura estructurada."
            )
        except (ValidationError, ValueError) as exc:
            raise GeminiResponseError(
                "Gemini devolvió datos con un formato inválido."
            ) from exc
        except GeminiReaderError:
            raise
        except Exception as exc:
            message = str(exc).upper()
            if "RESOURCE_EXHAUSTED" in message or "429" in message:
                safe_message = "La cuota de digitalización está temporalmente agotada."
            elif any(term in message for term in ("TIMEOUT", "NETWORK", "CONNECTION")):
                safe_message = "No pude contactar al servicio de digitalización."
            else:
                safe_message = "El servicio de digitalización devolvió un error."
            raise GeminiReaderError(safe_message) from exc
        finally:
            close = getattr(client, "close", None)
            if callable(close):
                close()
