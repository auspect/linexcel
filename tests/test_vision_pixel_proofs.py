"""Real decoded pixels constrain claims; provider doubles test calls, not OCR truth."""

import io
import json
import struct
import zlib

import pytest
from vision_helpers import image_bytes, region_response

from linexcel.aidoc import TokenUsage, describe_images
from linexcel.i18n import LANGUAGES
from linexcel.vision_evidence import (
    VisionEvidenceError,
    parse_observations,
    prepare_image,
)


def png(source):
    buffer = io.BytesIO()
    source.save(buffer, format="PNG")
    return buffer.getvalue()


def rgb16_png(width=40, height=30, *, alpha=False):
    """A valid native 16-bit PNG whose variation Pillow reduces to RGB8."""

    def chunk(name, data):
        return (
            struct.pack(">I", len(data))
            + name
            + data
            + struct.pack(">I", zlib.crc32(name + data))
        )

    rows = []
    for y in range(height):
        values = []
        for x in range(width):
            varied = (x, y) == (width // 2, height // 2)
            channels = [1001 if varied else 1000, 1000, 1000]
            if alpha:
                channels.append(1 if varied else 0)
            values.append(struct.pack(">" + "H" * len(channels), *channels))
        rows.append(b"\x00" + b"".join(values))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(
            b"IHDR",
            struct.pack(">IIBBBBB", width, height, 16, 6 if alpha else 2, 0, 0, 0),
        )
        + chunk(b"IDAT", zlib.compress(b"".join(rows)))
        + chunk(b"IEND", b"")
    )


@pytest.mark.parametrize(
    "mode,color",
    [
        ("RGB", "white"),
        ("RGB", "black"),
        ("RGB", "blue"),
        ("L", 128),
        ("RGBA", (1, 2, 3, 127)),
    ],
)
def test_exact_uniform_pixels_are_decided_without_model_or_configuration(mode, color):
    image = pytest.importorskip("PIL.Image")
    payload = png(image.new(mode, (1800, 1300), color))
    checks, usage = {}, TokenUsage()
    # No provider, model or API key: pixel-only decisions need no AI setup.
    described = describe_images(
        {"S": payload}, usage=usage, token_budget=0, validation_results=checks
    )
    page = checks["S"]["pages"][0]
    assert "4 exactly uniform" in described["S"]
    assert checks["S"]["status"] == "complete"
    assert usage.total == usage.requests == 0
    assert page["region_coverage_complete"]
    assert page["model_regions"] == []
    assert len(page["deterministic_regions"]) == 4
    assert page["completed_regions"] == page["deterministic_regions"]
    for batch in page["batches"]:
        assert batch["status"] == "complete"
        assert batch["resolution"] == "deterministic_pixels"
        assert batch["attempts"] == []
        assert batch["provider_request_required"] is False
        assert "provider_receipt" not in batch and "raw_response" not in batch
        assert batch["observations"][0]["observations"] == []
        assert batch["observations"][0]["observation_origin"] == "deterministic_pixels"


def test_transparency_is_decided_from_alpha_not_invisible_rgb_values():
    image = pytest.importorskip("PIL.Image")
    source = image.new("RGBA", (80, 80), (255, 255, 255, 0))
    source.putpixel((40, 40), (0, 0, 0, 0))
    native = prepare_image(png(source))
    assert native.regions[0].text_visibility == "none_fully_transparent"
    checks = {}
    describe_images({"S": png(source)}, validation_results=checks)
    assert checks["S"]["pages"][0]["model_regions"] == []
    # A single visible, almost transparent pixel must not be dropped.
    source.putpixel((40, 40), (0, 0, 0, 1))
    assert prepare_image(png(source)).regions[0].text_visibility == "not_determined"


@pytest.mark.parametrize("color", [(254, 254, 254), (255, 255, 254), (0, 0, 0)])
def test_one_low_contrast_or_dark_pixel_prevents_uniform_shortcut(color):
    image = pytest.importorskip("PIL.Image")
    source = image.new("RGB", (40, 30), "white")
    source.putpixel((20, 15), color)
    assert prepare_image(png(source)).regions[0].text_visibility == "not_determined"


def test_palette_indices_are_not_confused_with_visible_colors():
    image = pytest.importorskip("PIL.Image")
    source = image.new("P", (40, 30), 0)
    source.putpalette([255, 255, 255, 255, 255, 255] + [0] * 762)
    source.putpixel((20, 15), 1)
    assert prepare_image(png(source)).regions[0].text_visibility == "none_exact_uniform"
    source.putpalette([255, 255, 255, 254, 254, 254] + [0] * 762)
    assert prepare_image(png(source)).regions[0].text_visibility == "not_determined"


@pytest.mark.parametrize("value", [1001, 2000])
def test_high_depth_pixels_are_never_classified_using_a_lossy_rgba_conversion(value):
    image = pytest.importorskip("PIL.Image")
    source = image.new("I;16", (40, 30), 1000)
    source.putpixel((20, 15), value)
    payload = png(source)
    assert source.getextrema() == (1000, value)
    assert source.convert("RGBA").getextrema() == ((255, 255),) * 4
    native = prepare_image(payload)
    assert native.regions[0].payload == payload
    assert native.regions[0].text_visibility == "not_determined"
    proof = native.regions[0].pixel_evidence()
    assert proof["decoded_mode"] == "I;16"
    assert proof["rgba_extrema"] is None

    class Provider:
        def __init__(self):
            self.requests = 0

        def generate(self, *args, **kwargs):
            raise AssertionError("Pixels required")

        def generate_with_images(self, system, user, images, **kwargs):
            self.requests += 1
            assert images[0][1] == payload
            return json.dumps(
                {
                    "regions": [
                        {"id": "r0-c0", "readability": "unreadable", "observations": []}
                    ]
                }
            ), TokenUsage(20, 10, 1)

    provider, checks = Provider(), {}
    describe_images({"S": payload}, provider=provider, validation_results=checks)
    assert provider.requests == 1
    assert checks["S"]["pages"][0]["deterministic_regions"] == []


@pytest.mark.parametrize("alpha", [False, True])
def test_rgb16_header_prevents_shortcut_even_when_pillow_decodes_to_rgb8(alpha):
    image = pytest.importorskip("PIL.Image")
    payload = rgb16_png(alpha=alpha)
    assert payload[24] == 16
    with image.open(io.BytesIO(payload)) as source:
        assert source.mode == ("RGBA" if alpha else "RGB")
        assert all(low == high for low, high in source.getextrema())
    native = prepare_image(payload)
    assert native.metadata()["encoded_sample_depth"] == 16
    assert native.metadata()["source_format"] == "PNG"
    assert native.regions[0].payload == payload
    proof = native.regions[0].pixel_evidence()
    assert proof["decoded_mode"] == ("RGBA" if alpha else "RGB")
    assert proof["source_precision_safe"] is False
    assert proof["rgba_extrema"] is None
    assert proof["text_visibility"] == "not_determined"

    class Provider:
        def __init__(self):
            self.requests = 0

        def generate(self, *args, **kwargs):
            raise AssertionError("Pixels required")

        def generate_with_images(self, system, user, images, **kwargs):
            self.requests += 1
            assert images[0][1] == payload
            return json.dumps(
                {
                    "regions": [
                        {"id": "r0-c0", "readability": "unreadable", "observations": []}
                    ]
                }
            ), TokenUsage(20, 10, 1)

    provider, checks = Provider(), {}
    describe_images({"S": payload}, provider=provider, validation_results=checks)
    assert provider.requests == 1
    assert checks["S"]["pages"][0]["deterministic_regions"] == []


def test_rgb16_that_cannot_be_cropped_losslessly_fails_explicitly():
    with pytest.raises(VisionEvidenceError, match="encoded sample precision"):
        prepare_image(rgb16_png(width=1025, height=2))


def test_grayscale16_large_native_crops_keep_their_encoded_depth():
    image = pytest.importorskip("PIL.Image")
    source = image.new("I;16", (1025, 30), 1000)
    source.putpixel((1000, 15), 2000)
    native = prepare_image(png(source))
    assert len(native.regions) == 2
    for region in native.regions:
        assert region.payload[24] == 16
        assert region.text_visibility == "not_determined"
        with image.open(io.BytesIO(region.payload)) as cropped:
            assert cropped.tobytes() == source.crop(region.bbox).tobytes()


@pytest.mark.parametrize("readability", ["readable", "partly_readable", "unreadable"])
def test_uniform_pixels_reject_a_conflicting_model_readability(readability):
    image = pytest.importorskip("PIL.Image")
    native = prepare_image(png(image.new("RGB", (40, 30), "white")))
    reply = json.dumps(
        {"regions": [{"id": "r0-c0", "readability": readability, "observations": []}]}
    )
    with pytest.raises(VisionEvidenceError, match="uniform"):
        parse_observations(reply, native, list(native.regions))


def test_old_white_crop_hallucination_is_rejected_not_merely_relabelled():
    image = pytest.importorskip("PIL.Image")
    native = prepare_image(png(image.new("RGB", (710, 1024), "white")))
    prompt = json.dumps({"regions": [r.metadata() for r in native.regions]})
    # Reproduces the well-formed but false 12 October Sales response.
    reply = json.loads(region_response(prompt, "Chris O'Connor"))
    reply["regions"][0]["observations"].extend(
        [
            {"text": "United States", "certainty": "visible"},
            {"text": "23", "certainty": "visible"},
        ]
    )
    with pytest.raises(VisionEvidenceError, match="uniform"):
        parse_observations(json.dumps(reply), native, list(native.regions))


def test_mixed_capture_keeps_all_pixels_but_requests_only_nonuniform_crop(tmp_path):
    image = pytest.importorskip("PIL.Image")
    draw = pytest.importorskip("PIL.ImageDraw")
    source = image.new("RGB", (1800, 1300), "white")
    draw.Draw(source).text((1250, 1150), "Flag", fill="black")
    payload = png(source)

    class Provider:
        def __init__(self):
            self.calls = []

        def generate(self, *args, **kwargs):
            raise AssertionError("A nonuniform capture requires pixels")

        def generate_with_images(self, system, user, images, **kwargs):
            self.calls.append((user, images, kwargs))
            return region_response(user, "Flag"), TokenUsage(20, 10, 1)

    provider, checks, usage = Provider(), {}, TokenUsage()
    described = describe_images(
        {"S": payload},
        provider=provider,
        usage=usage,
        validation_results=checks,
        evidence_dir=tmp_path,
    )
    page = checks["S"]["pages"][0]
    assert len(page["evidence"]["regions"]) == 4
    assert len(page["deterministic_regions"]) == 3
    assert page["model_regions"] == ["r1-c1"]
    assert len(provider.calls) == usage.requests == 1
    prompt, sent, options = provider.calls[0]
    assert options["temperature"] == 0
    assert [x["id"] for x in json.loads(prompt)["regions"]] == ["r1-c1"]
    native = prepare_image(payload)
    assert sent[0][1] == native.regions[-1].payload
    assert "“Flag”" in described["S"]
    observation = next(o for o in page["observations"] if "text" in o)
    assert observation["certainty_source"] == "model_self_report_not_pixel_verification"
    assert observation["factual_certification"] is False


@pytest.mark.parametrize("language", sorted(LANGUAGES))
def test_local_pixel_abstention_is_localized_and_never_certifies_model_text(language):
    image = pytest.importorskip("PIL.Image")
    checks = {}
    docs = describe_images(
        {"S": png(image.new("RGB", (40, 30), "white"))},
        language=language,
        validation_results=checks,
    )
    assert docs["S"].strip()
    assert not checks["S"]["factual_certification"]
    assert "“" not in docs["S"]


def test_nonuniform_pixels_never_verify_a_numerical_transcription():
    native = prepare_image(image_bytes())
    prompt = json.dumps({"regions": [r.metadata() for r in native.regions]})
    observation = parse_observations(
        region_response(prompt, "50693,75"), native, list(native.regions)
    )[0]
    # This layer cannot prove arbitrary OCR correct. Keep this limitation typed,
    # even when a model calls an unsupported number visible.
    assert observation["certainty"] == "visible"
    assert observation["certainty_source"] == "model_self_report_not_pixel_verification"
    assert observation["factual_certification"] is False
