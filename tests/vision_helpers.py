import io
import json

import pytest


def image_bytes(width=40, height=30):
    """Nonuniform pixels for provider-contract tests, not an OCR truth oracle."""
    image = pytest.importorskip("PIL.Image")
    buffer = io.BytesIO()
    source = image.new("RGB", (width, height), "white")
    for x in range(0, width, 97):
        source.putpixel((x, height // 2), (23, 49, 82))
    source.save(buffer, format="PNG")
    return buffer.getvalue()


def region_response(prompt, text="Readable header"):
    regions = json.loads(prompt)["regions"]
    return json.dumps(
        {
            "regions": [
                {
                    "id": region["id"],
                    "readability": "readable",
                    "observations": [
                        {
                            "text": text,
                            "certainty": "visible",
                        }
                    ],
                }
                for region in regions
            ]
        }
    )
