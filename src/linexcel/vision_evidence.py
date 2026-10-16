"""Native screenshot regions and bounded, reviewable visual observations.

Coordinates are pixels in the decoded original image, with exclusive right
and bottom edges. Cropping never rescales pixels. Valid JSON and coordinates
establish traceability, not the factual truth of a model's observation.
"""

from __future__ import annotations

import hashlib
import html
import io
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

TILE_SIZE = 1024
TILE_OVERLAP = 128
BATCH_SIZE = 1
MAX_OBSERVATIONS = 3
MAX_OBSERVATION_CHARS = 80
PROTOCOL = "linexcel-native-vision-v4"
_RGBA_SAFE_MODES = {"1", "L", "LA", "P", "RGB", "RGBA"}


class VisionEvidenceError(ValueError):
    """Pixels or model output cannot satisfy the visual evidence contract."""


@dataclass(frozen=True)
class Region:
    id: str
    bbox: tuple[int, int, int, int]
    payload: bytes
    media_type: str
    rgba_extrema: tuple[tuple[int, int], ...] | None
    decoded_mode: str
    source_precision_safe: bool

    @property
    def text_visibility(self) -> str:
        """Rule out text only when exact decoded pixels make it impossible."""
        if self.rgba_extrema is None:
            return "not_determined"
        if self.rgba_extrema[3] == (0, 0):
            return "none_fully_transparent"
        if all(low == high for low, high in self.rgba_extrema):
            return "none_exact_uniform"
        return "not_determined"

    def pixel_evidence(self) -> dict[str, Any]:
        return {
            "method": "exact_decoded_rgba_extrema"
            if self.rgba_extrema is not None
            else "not_analyzed_unsupported_sample_precision_or_mode",
            "decoded_mode": self.decoded_mode,
            "source_precision_safe": self.source_precision_safe,
            "rgba_extrema": [list(channel) for channel in self.rgba_extrema]
            if self.rgba_extrema is not None
            else None,
            "text_visibility": self.text_visibility,
            "scope": "this_native_crop_only_not_a_blank_sheet_claim",
        }

    @property
    def sha256(self) -> str:
        return hashlib.sha256(self.payload).hexdigest()

    def metadata(self) -> dict[str, Any]:
        x0, y0, x1, y1 = self.bbox
        return {
            "id": self.id,
            "bbox": list(self.bbox),
            "width": x1 - x0,
            "height": y1 - y0,
            "pixel_scale": 1,
            "scope": "entire_native_crop",
            "sha256": self.sha256,
            "media_type": self.media_type,
            "pixel_evidence": self.pixel_evidence(),
        }


@dataclass(frozen=True)
class ImageEvidence:
    sha256: str
    width: int
    height: int
    regions: tuple[Region, ...]
    decoder_version: str
    source_format: str
    encoded_sample_depth: int | None

    def metadata(self) -> dict[str, Any]:
        return {
            "protocol": PROTOCOL,
            "original_sha256": self.sha256,
            "width": self.width,
            "height": self.height,
            "tile_size": TILE_SIZE,
            "overlap": TILE_OVERLAP,
            "decoder": "Pillow",
            "decoder_version": self.decoder_version,
            "source_format": self.source_format,
            "encoded_sample_depth": self.encoded_sample_depth,
            "pixel_coverage_complete": True,
            "regions": [region.metadata() for region in self.regions],
            "factual_certification": False,
        }


def _starts(length: int) -> list[int]:
    positions = [0]
    while positions[-1] + TILE_SIZE < length:
        positions.append(positions[-1] + TILE_SIZE - TILE_OVERLAP)
    return positions


def _encoded_sample_depth(
    payload: bytes, source_format: str, jpeg_bits: Any
) -> int | None:
    """Keep source precision separate from Pillow's potentially reduced mode."""
    if source_format == "PNG":
        # IHDR is the first PNG chunk, with a fixed 13-byte body. Pillow still
        # validates/decodes the image; unknown encodings cannot enable a shortcut.
        if (
            len(payload) >= 33
            and payload[:8] == b"\x89PNG\r\n\x1a\n"
            and payload[8:16] == b"\x00\x00\x00\rIHDR"
            and payload[24] in {1, 2, 4, 8, 16}
        ):
            return payload[24]
        return None
    if source_format == "JPEG":
        return jpeg_bits if type(jpeg_bits) is int else None
    if source_format in {"GIF", "WEBP"}:
        return 8
    return None


def prepare_image(payload: bytes) -> ImageEvidence:
    """Decode and cover the entire image with deterministic native regions."""
    try:
        from PIL import Image, __version__
    except ImportError as exc:
        raise VisionEvidenceError(
            "Native screenshot evidence needs Pillow; install linexcel[screenshots]."
        ) from exc
    original_sha = hashlib.sha256(payload).hexdigest()
    try:
        with Image.open(io.BytesIO(payload)) as source:
            if getattr(source, "n_frames", 1) != 1:
                raise VisionEvidenceError("Animated images need separate frame images")
            source_format = source.format or ""
            sample_depth = _encoded_sample_depth(
                payload, source_format, getattr(source, "bits", None)
            )
            precision_safe = sample_depth is not None and sample_depth <= 8
            source.load()
            width, height = source.size
            regions = []
            for row, y in enumerate(_starts(height)):
                for column, x in enumerate(_starts(width)):
                    bbox = (x, y, min(x + TILE_SIZE, width), min(y + TILE_SIZE, height))
                    crop = source.crop(bbox)
                    extrema = None
                    if precision_safe and crop.mode in _RGBA_SAFE_MODES:
                        # Converting I;16/I/F to RGBA can saturate different
                        # native values to the same 8-bit color. Such a lossy
                        # conversion must never establish a uniform crop.
                        rgba = crop.convert("RGBA")
                        channel_extrema: list[tuple[int, int]] = []
                        for channel in ("R", "G", "B", "A"):
                            low, high = rgba.getchannel(channel).getextrema()
                            if not isinstance(low, int) or not isinstance(high, int):
                                raise VisionEvidenceError(
                                    "Expected 8-bit RGBA channels"
                                )
                            channel_extrema.append((low, high))
                        extrema = tuple(channel_extrema)
                    if bbox == (0, 0, width, height) and source.format in {
                        "PNG",
                        "JPEG",
                        "WEBP",
                        "GIF",
                    }:
                        data = payload
                        media_type = {
                            "PNG": "image/png",
                            "JPEG": "image/jpeg",
                            "WEBP": "image/webp",
                            "GIF": "image/gif",
                        }[source.format]
                    else:
                        if (
                            sample_depth is not None
                            and sample_depth > 8
                            and crop.mode in _RGBA_SAFE_MODES
                        ):
                            raise VisionEvidenceError(
                                "High-depth image cannot be cropped without losing "
                                "encoded sample precision"
                            )
                        output = io.BytesIO()
                        crop.save(output, format="PNG")
                        data, media_type = output.getvalue(), "image/png"
                    regions.append(
                        Region(
                            f"r{row}-c{column}",
                            bbox,
                            data,
                            media_type,
                            extrema,
                            crop.mode,
                            precision_safe,
                        )
                    )
    except (OSError, SyntaxError, Image.DecompressionBombError) as exc:
        raise VisionEvidenceError(f"Cannot decode screenshot: {exc}") from exc
    return ImageEvidence(
        original_sha,
        width,
        height,
        tuple(regions),
        __version__,
        source_format,
        sample_depth,
    )


def persist_image(evidence: ImageEvidence, folder: Path) -> dict[str, Any]:
    """Save exact sent bytes, refusing to replace a different crop."""
    target = folder / evidence.sha256
    target.mkdir(parents=True, exist_ok=True)
    metadata = evidence.metadata()
    for region, record in zip(evidence.regions, metadata["regions"], strict=True):
        suffix = {
            "image/png": ".png",
            "image/jpeg": ".jpg",
            "image/webp": ".webp",
            "image/gif": ".gif",
        }[region.media_type]
        path = target / f"{region.id}-{region.sha256[:16]}{suffix}"
        if path.exists() and path.read_bytes() != region.payload:
            raise VisionEvidenceError(f"Refusing to replace different pixels: {path}")
        if not path.exists():
            path.write_bytes(region.payload)
        record["path"] = str(path.resolve())
    path = target / "manifest.json"
    # A decoder/version change may produce different PNG bytes for unchanged
    # pixels. Do not silently overwrite the earlier sent-byte manifest.
    content = json.dumps(metadata, ensure_ascii=False, indent=2) + "\n"
    if path.exists() and path.read_text(encoding="utf-8") != content:
        raise VisionEvidenceError(f"Existing crop manifest differs: {path}")
    path.write_text(content, encoding="utf-8")
    metadata["manifest_path"] = str(path.resolve())
    return metadata


def batch_prompt(
    name: str, page: int, evidence: ImageEvidence, regions: list[Region]
) -> str:
    return json.dumps(
        {
            "protocol": PROTOCOL,
            "sheet_label_untrusted": name,
            "page": page,
            "original": {
                "sha256": evidence.sha256,
                "width": evidence.width,
                "height": evidence.height,
            },
            "regions": [region.metadata() for region in regions],
            "coordinate_semantics": (
                "Observation scope is the entire native crop, computed by code. "
                "The model must not return coordinates or infer a glyph box. "
                "Crop bbox=[left,top,right,bottom] in original pixels; "
                "exclusive right/bottom."
            ),
            "scope": (
                "Only these native crops. Other crops and workbook data "
                "are not evidence in this request."
            ),
            "confidence_scope": (
                "certainty is a model self-report, not verification. "
                "Pixel extrema can rule out text in exactly uniform or fully "
                "transparent crops; nonuniform pixels do not verify any text. "
                "Prefer abstention to inferred labels or reconstructed numbers."
            ),
        },
        ensure_ascii=False,
        separators=(",", ":"),
    )


def parse_observations(
    text: str, evidence: ImageEvidence, regions: list[Region]
) -> list[dict[str, Any]]:
    """Accept short local quotations; derive their entire crop scope in code."""
    stripped = text.strip()
    if stripped.startswith("```json\n") and stripped.endswith("\n```"):
        stripped = stripped[8:-4]
    try:
        response = json.loads(stripped)
    except (json.JSONDecodeError, TypeError) as exc:
        raise VisionEvidenceError(
            "Vision response must be the region JSON contract"
        ) from exc
    if not isinstance(response, dict) or set(response) != {"regions"}:
        raise VisionEvidenceError("Vision response requires exactly a regions list")
    records = response["regions"]
    expected = {region.id: region for region in regions}
    if not isinstance(records, list) or len(records) != len(expected):
        raise VisionEvidenceError("Vision response omits or adds regions")
    seen = set()
    observations = []
    for record in records:
        if not isinstance(record, dict) or set(record) != {
            "id",
            "observations",
            "readability",
        }:
            raise VisionEvidenceError("Invalid vision region schema")
        identifier = record["id"]
        if (
            not isinstance(identifier, str)
            or identifier not in expected
            or identifier in seen
        ):
            raise VisionEvidenceError("Unknown or duplicate vision region")
        seen.add(identifier)
        region = expected[identifier]
        if not isinstance(record["readability"], str) or record["readability"] not in {
            "readable",
            "partly_readable",
            "unreadable",
            "no_visible_content",
        }:
            raise VisionEvidenceError("Invalid region readability")
        items = record["observations"]
        if not isinstance(items, list) or len(items) > MAX_OBSERVATIONS:
            raise VisionEvidenceError("Too many regional observations")
        if region.text_visibility != "not_determined" and (
            record["readability"] != "no_visible_content" or items
        ):
            raise VisionEvidenceError(
                "Exact uniform or fully transparent pixels cannot support text"
            )
        if record["readability"] in {"unreadable", "no_visible_content"} and items:
            raise VisionEvidenceError(
                "Unreadable or empty region cannot have factual observations"
            )
        if not items:
            observations.append(
                {
                    "region_id": identifier,
                    "readability": record["readability"],
                    "observations": [],
                    "observation_origin": "model",
                }
            )
        for item in items:
            if not isinstance(item, dict) or set(item) != {"text", "certainty"}:
                raise VisionEvidenceError("Invalid visual observation schema")
            if not isinstance(item["certainty"], str) or item["certainty"] not in {
                "visible",
                "ambiguous",
            }:
                raise VisionEvidenceError("Invalid visual observation certainty")
            rw, rh = region.bbox[2] - region.bbox[0], region.bbox[3] - region.bbox[1]
            label = item["text"]
            if (
                not isinstance(label, str)
                or not label.strip()
                or len(label) > MAX_OBSERVATION_CHARS
                or any(ord(character) < 32 for character in label)
            ):
                raise VisionEvidenceError(
                    "Visual observation text is empty or too long"
                )
            observations.append(
                {
                    **item,
                    "kind": "text",
                    "bbox": [0, 0, rw, rh],
                    "scope": "entire_native_crop",
                    "region_id": identifier,
                    "readability": record["readability"],
                    "original_bbox": list(region.bbox),
                    "original_sha256": evidence.sha256,
                    "crop_sha256": region.sha256,
                    "observation_origin": "model",
                    "certainty_source": "model_self_report_not_pixel_verification",
                    "factual_certification": False,
                }
            )
    return observations


def uniform_region_observations(
    regions: list[Region],
) -> list[dict[str, Any]]:
    """Describe exact no-text pixels without asking a model to infer content."""
    if not regions or any(r.text_visibility == "not_determined" for r in regions):
        raise VisionEvidenceError("A nonuniform crop requires visual inspection")
    return [
        {
            "region_id": region.id,
            "readability": "no_visible_content",
            "observations": [],
            "observation_origin": "deterministic_pixels",
            "pixel_evidence": region.pixel_evidence(),
            "original_bbox": list(region.bbox),
            "crop_sha256": region.sha256,
        }
        for region in regions
    ]


_LABELS = {
    "en": (
        "Native visual evidence",
        "image(s)",
        "region(s)",
        "Pixel scale 1; coordinates refer to each original image. "
        "AI observations remain unverified; "
        "pixels do not establish workbook calculations.",
        "No confidently transcribed observation. Inspect the native regions.",
        "Additional bounded visual observations",
    ),
    "fr": (
        "Preuves visuelles natives",
        "image(s)",
        "région(s)",
        "Échelle pixel 1 ; coordonnées de chaque image originale. "
        "Les observations IA restent non vérifiées ; "
        "les pixels ne prouvent aucun calcul du classeur.",
        "Aucune observation transcrite avec confiance. Inspecter les régions natives.",
        "Autres observations visuelles bornées",
    ),
    "de": (
        "Native visuelle Belege",
        "Bild(er)",
        "Region(en)",
        "Pixelskala 1; Koordinaten beziehen sich auf das Originalbild. "
        "KI-Beobachtungen bleiben ungeprüft; "
        "Pixel belegen keine Arbeitsmappenberechnung.",
        "Keine zuverlässig transkribierte Beobachtung. Native Regionen prüfen.",
        "Weitere räumlich begrenzte Beobachtungen",
    ),
    "es": (
        "Evidencia visual nativa",
        "imagen(es)",
        "región(es)",
        "Escala de píxel 1; coordenadas de cada imagen original. "
        "Las observaciones IA no están verificadas; "
        "los píxeles no prueban cálculos del libro.",
        "Ninguna observación transcrita con confianza. Revise las regiones nativas.",
        "Otras observaciones visuales acotadas",
    ),
    "it": (
        "Evidenza visiva nativa",
        "immagine/i",
        "regione/i",
        "Scala pixel 1; coordinate di ogni immagine originale. "
        "Le osservazioni IA restano non verificate; "
        "i pixel non dimostrano calcoli del libro.",
        "Nessuna osservazione trascritta con sicurezza. Esaminare le regioni native.",
        "Altre osservazioni visive delimitate",
    ),
    "ja": (
        "元の解像度の視覚的証拠",
        "画像",
        "領域",
        "ピクセル倍率1。座標は各元画像を参照します。AI観察は未検証で、ピクセルだけではブックの計算を証明できません。",
        "確実に転記できた観察はありません。元の領域を確認してください。",
        "その他の領域限定の視覚的観察",
    ),
    "nl": (
        "Visueel bewijs op oorspronkelijke resolutie",
        "afbeelding(en)",
        "regio('s)",
        "Pixelschaal 1; coördinaten verwijzen naar elke originele afbeelding. "
        "AI-waarnemingen zijn niet geverifieerd; "
        "pixels bewijzen geen werkmapberekening.",
        "Geen betrouwbaar getranscribeerde waarneming. "
        "Bekijk de oorspronkelijke regio's.",
        "Overige begrensde visuele waarnemingen",
    ),
    "pt": (
        "Evidência visual nativa",
        "imagem(ns)",
        "região(ões)",
        "Escala de pixel 1; coordenadas de cada imagem original. "
        "As observações IA não estão verificadas; "
        "pixels não provam cálculos da pasta de trabalho.",
        "Nenhuma observação transcrita com confiança. Inspecione as regiões nativas.",
        "Outras observações visuais delimitadas",
    ),
    "zh": (
        "原始分辨率视觉证据",
        "图像",
        "区域",
        "像素比例1；坐标对应每张原始图像。AI观察尚未验证，像素不能证明工作簿计算。",
        "没有可可靠转录的观察。请检查原始区域。",
        "其他限定区域的视觉观察",
    ),
}

_PIXEL_LABELS = {
    "en": (
        "Local pixel check: {count} exactly uniform or fully transparent crop(s); "
        "no visible text. No AI request for these crops."
    ),
    "fr": (
        "Contrôle local des pixels : {count} région(s) exactement uniforme(s) "
        "ou entièrement transparente(s), sans texte visible. "
        "Aucun appel IA pour ces régions."
    ),
    "de": (
        "Lokale Pixelprüfung: {count} exakt einfarbige oder vollständig "
        "transparente Bereiche ohne sichtbaren Text. "
        "Keine KI-Anfrage für diese Bereiche."
    ),
    "es": (
        "Comprobación local de píxeles: {count} regiones exactamente uniformes "
        "o completamente transparentes, sin texto visible. "
        "Sin solicitud de IA para estas regiones."
    ),
    "it": (
        "Controllo locale dei pixel: {count} regioni esattamente uniformi "
        "o completamente trasparenti, senza testo visibile. "
        "Nessuna richiesta IA per queste regioni."
    ),
    "ja": (
        "ローカルのピクセル検査：完全に均一または完全に透明な領域が{count}個あり、"
        "表示される文字はありません。これらの領域についてAIへのリクエストは行っていません。"
    ),
    "nl": (
        "Lokale pixelcontrole: {count} exact egale of volledig transparante "
        "gebieden zonder zichtbare tekst. Geen AI-aanvraag voor deze gebieden."
    ),
    "pt": (
        "Verificação local dos píxeis: {count} regiões exatamente uniformes "
        "ou totalmente transparentes, sem texto visível. "
        "Sem pedido de IA para estas regiões."
    ),
    "zh": (
        "本地像素检查：{count}个区域完全均匀或完全透明，没有可见文字。"
        "这些区域未发送AI请求。"
    ),
}


def render_observations(pages: list[dict[str, Any]], language: str = "en") -> str:
    """Bound the initial view, retain every regional observation in a disclosure."""
    labels = []
    for page_number, page in enumerate(pages, 1):
        for observation in page["observations"]:
            if "text" not in observation:
                continue
            text = html.escape(observation["text"].strip())
            text = re.sub(r"([\\`*_{}\[\]()#+!|])", r"\\\1", text)
            if observation["kind"] == "text":
                text = f"“{text}”"
            certainty = observation["certainty"]
            label = (
                f"- {text} — `{certainty}` · p{page_number}/"
                f"{observation['region_id']} · pixels "
                f"`{observation['original_bbox']}` · crop "
                f"`{observation['crop_sha256'][:12]}`"
            )
            if label not in labels:
                labels.append(label)
    total = sum(len(page["evidence"]["regions"]) for page in pages)
    title, image_label, region_label, disclaimer, empty_label, more_label = _LABELS[
        language
    ]
    result = (
        f"**{title}:** {len(pages)} {image_label}, {total} {region_label}. "
        f"{disclaimer}\n\n"
    )
    deterministic_count = sum(
        observation.get("observation_origin") == "deterministic_pixels"
        for page in pages
        for observation in page["observations"]
    )
    if deterministic_count:
        result += _PIXEL_LABELS[language].format(count=deterministic_count) + "\n\n"
    if not labels:
        return result + empty_label
    # Spread visible examples across the entire processed source; do not infer
    # any global summary from these local observations.
    indices = (
        sorted(
            {
                round(i * (len(labels) - 1) / min(5, len(labels) - 1))
                for i in range(min(6, len(labels)))
            }
        )
        if len(labels) > 1
        else [0]
    )
    result += "\n".join(labels[index] for index in indices)
    remaining = [label for index, label in enumerate(labels) if index not in indices]
    if remaining:
        result += (
            f"\n\n<details><summary>{more_label} "
            f"({len(remaining)})</summary>\n\n"
            + "\n".join(remaining)
            + "\n\n</details>"
        )
    return result
