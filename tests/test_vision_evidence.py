"""Native crop coverage and response scope; providers are local test doubles."""

import hashlib
import io
import json

import pytest
from vision_helpers import image_bytes, region_response

from linexcel.aidoc import AiDocError, TokenUsage, describe_images
from linexcel.i18n import LANGUAGES
from linexcel.vision_evidence import (
    BATCH_SIZE,
    VisionEvidenceError,
    batch_prompt,
    parse_observations,
    persist_image,
    prepare_image,
)


class NativeProvider:
    def __init__(self):
        self.calls = []

    def generate(self, *args, **kwargs):
        raise AssertionError("Vision must include pixels")

    def generate_with_images(self, system, user, images, **kwargs):
        self.calls.append((system, user, images))
        return region_response(user), TokenUsage(20, 10, 1)


@pytest.mark.parametrize(
    "size", [(40, 30), (1024, 1024), (1025, 1025), (18574, 642), (1564, 10417)]
)
def test_tiles_cover_every_original_pixel_without_resampling(size):
    image_module = pytest.importorskip("PIL.Image")
    # A nonuniform source makes off-by-one crop coordinates observable.
    source = image_module.new("RGB", size, "white")
    for x in range(0, size[0], 97):
        source.putpixel((x, size[1] - 1), (x % 256, 23, 99))
    buffer = io.BytesIO()
    source.save(buffer, format="PNG")
    native = prepare_image(buffer.getvalue())
    reconstructed = image_module.new("RGB", size, "black")
    coverage = image_module.new("1", size)
    for region in native.regions:
        left, top, right, bottom = region.bbox
        assert right - left <= 1024 and bottom - top <= 1024
        with image_module.open(io.BytesIO(region.payload)) as crop:
            assert crop.size == (right - left, bottom - top)
            assert crop.tobytes() == source.crop(region.bbox).tobytes()
            reconstructed.paste(crop, (left, top))
        coverage.paste(1, region.bbox)
    assert coverage.getextrema() == (1, 1)
    assert reconstructed.tobytes() == source.tobytes()
    assert native.metadata()["pixel_coverage_complete"]
    assert all(r["pixel_scale"] == 1 for r in native.metadata()["regions"])


def test_small_capture_preserves_exact_sent_original_bytes():
    payload = image_bytes(895, 257)
    (region,) = prepare_image(payload).regions
    assert region.payload == payload
    assert region.bbox == (0, 0, 895, 257)


def test_crop_manifest_is_stable_and_refuses_replacement(tmp_path):
    native = prepare_image(image_bytes(2050, 30))
    first = persist_image(native, tmp_path)
    assert first == persist_image(native, tmp_path)
    from pathlib import Path

    crop = Path(first["regions"][0]["path"])
    assert hashlib.sha256(crop.read_bytes()).hexdigest() == native.regions[0].sha256
    crop.write_bytes(b"different pixels")
    with pytest.raises(VisionEvidenceError, match="replace different"):
        persist_image(native, tmp_path)


def response_for(native, **overrides):
    region = native.regions[0]
    item = {
        "text": "£0",
        "certainty": "visible",
    }
    item.update(overrides)
    return json.dumps(
        {
            "regions": [
                {"id": region.id, "readability": "readable", "observations": [item]}
            ]
        }
    )


@pytest.mark.parametrize(
    "box",
    [
        [-1, 0, 10, 10],
        [0, 0, 41, 10],
        [0, 0, 10, 31],
        [0, 0, 0, 10],
        [False, 0, 10, 10],
        [0, 0, 1.5, 10],
    ],
)
def test_model_cannot_supply_its_own_geometry(box):
    native = prepare_image(image_bytes())
    with pytest.raises(VisionEvidenceError, match="schema"):
        parse_observations(response_for(native, bbox=box), native, list(native.regions))


@pytest.mark.parametrize(
    "text",
    [
        "Readable prose without boxes",
        '{"regions":[]}',
        '{"regions":[],"summary":"whole workbook"}',
    ],
)
def test_unbounded_or_incomplete_response_is_rejected(text):
    native = prepare_image(image_bytes())
    with pytest.raises(VisionEvidenceError):
        parse_observations(text, native, list(native.regions))


def test_model_cannot_add_a_clipping_claim_at_a_crop_border():
    native = prepare_image(image_bytes(2200, 40))
    middle = native.regions[1]
    reply = json.dumps(
        {
            "regions": [
                {
                    "id": middle.id,
                    "readability": "readable",
                    "observations": [
                        {
                            "kind": "clipping",
                            "bbox": [0, 5, 10, 15],
                            "text": "Possibly cut glyph",
                            "certainty": "ambiguous",
                        }
                    ],
                }
            ]
        }
    )
    with pytest.raises(VisionEvidenceError, match="schema"):
        parse_observations(reply, native, [middle])


def test_model_cannot_add_clipping_even_at_an_original_border():
    native = prepare_image(image_bytes())
    with pytest.raises(VisionEvidenceError, match="schema"):
        parse_observations(
            response_for(native, kind="clipping"), native, list(native.regions)
        )
    with pytest.raises(VisionEvidenceError, match="schema"):
        parse_observations(
            response_for(native, kind="clipping", certainty="ambiguous"),
            native,
            list(native.regions),
        )


def test_program_computes_entire_crop_scope_and_original_coordinates():
    native = prepare_image(image_bytes(2200, 40))
    region = native.regions[1]
    prompt = batch_prompt("Untrusted label", 1, native, [region])
    observations = parse_observations(region_response(prompt), native, [region])
    assert observations[0]["original_bbox"] == list(region.bbox)
    assert observations[0]["bbox"] == [0, 0, 1024, 40]
    assert observations[0]["scope"] == "entire_native_crop"
    assert observations[0]["original_sha256"] == native.sha256
    assert observations[0]["crop_sha256"] == region.sha256


@pytest.mark.parametrize("language", sorted(LANGUAGES))
def test_all_languages_preserve_region_contract_and_bounded_render(language):
    provider = NativeProvider()
    checks = {}
    docs = describe_images(
        {"Sheet": image_bytes(4600, 30)},
        provider=provider,
        language=language,
        validation_results=checks,
    )
    assert len(provider.calls) == 5
    assert all(len(call[2]) == 1 for call in provider.calls)
    assert "r0-c0" in docs["Sheet"]
    assert checks["Sheet"]["status"] == "complete"
    assert "LOCAL crop pixels" in provider.calls[0][0]
    assert not checks["Sheet"]["factual_certification"]


def test_large_images_and_all_pages_are_batched_and_traced(tmp_path):
    provider = NativeProvider()
    checks = {}
    first, second = image_bytes(4600, 30), image_bytes(40, 30)
    docs = describe_images(
        {"Sheet": [first, second]},
        provider=provider,
        validation_results=checks,
        evidence_dir=tmp_path,
    )
    pages = checks["Sheet"]["pages"]
    assert len(pages) == 2 and len(provider.calls) == 6
    assert sum(len(call[2]) for call in provider.calls) == 6
    assert all(len(call[2]) <= BATCH_SIZE for call in provider.calls)
    assert all(page["region_coverage_complete"] for page in pages)
    assert pages[0]["evidence"]["original_sha256"] == hashlib.sha256(first).hexdigest()
    assert pages[1]["evidence"]["original_sha256"] == hashlib.sha256(second).hexdigest()
    for _, prompt, images in provider.calls:
        metadata = json.loads(prompt)
        assert [(r["id"], r["sha256"]) for r in metadata["regions"]] == [
            (name, hashlib.sha256(data).hexdigest()) for name, data, _ in images
        ]
    assert "p2/r0-c0" in docs["Sheet"]


def test_failure_after_first_batch_retains_raw_pixels_but_not_complete_sheet():
    class FailedProvider(NativeProvider):
        def generate_with_images(self, system, user, images, **kwargs):
            if self.calls:
                self.calls.append((system, user, images))
                return '{"regions":[]}', TokenUsage(20, 10, 1)
            return super().generate_with_images(system, user, images, **kwargs)

    checks, usage = {}, TokenUsage()
    with pytest.raises(AiDocError):
        describe_images(
            {"Sheet": image_bytes(4600, 30)},
            provider=FailedProvider(),
            validation_results=checks,
            usage=usage,
        )
    report = checks["Sheet"]
    assert report["status"] == "incomplete"
    assert report["pages"][0]["batches"][0]["status"] == "complete"
    assert report["pages"][0]["batches"][1]["status"] == "invalid_response"
    assert report["pages"][0]["batches"][1]["raw_response"] == '{"regions":[]}'
    assert usage.total == 90
    assert [
        attempt["status"] for attempt in report["pages"][0]["batches"][1]["attempts"]
    ] == [
        "invalid_response",
        "invalid_response",
    ]


def test_budget_is_cumulative_even_without_external_usage_accumulator():
    provider = NativeProvider()
    with pytest.raises(AiDocError, match="budget"):
        describe_images(
            {"Sheet": image_bytes(4600, 30)}, provider=provider, token_budget=25
        )
    assert len(provider.calls) == 1


@pytest.mark.parametrize("width,budget,calls", [(40, 25, 1), (4600, 50, 2)])
def test_last_batch_cannot_complete_above_budget(width, budget, calls):
    provider = NativeProvider()
    checks, usage = {}, TokenUsage()
    with pytest.raises(AiDocError, match="budget") as caught:
        describe_images(
            {"Sheet": image_bytes(width, 30)},
            provider=provider,
            token_budget=budget,
            usage=usage,
            validation_results=checks,
        )
    assert len(provider.calls) == calls
    assert caught.value.usage.total == usage.total == calls * 30
    assert checks["Sheet"]["status"] == "incomplete"
    batch = checks["Sheet"]["pages"][0]["batches"][-1]
    assert batch["failure_code"] == "token_budget"
    assert batch["status"] == "failed"
    assert batch["raw_response"]


def test_exact_budget_can_complete_the_final_batch():
    checks = {}
    docs = describe_images(
        {"Sheet": image_bytes()},
        provider=NativeProvider(),
        token_budget=30,
        validation_results=checks,
    )
    assert "Sheet" in docs and checks["Sheet"]["status"] == "complete"


def test_untrusted_pixels_do_not_become_rendered_markup():
    class HostileProvider(NativeProvider):
        def generate_with_images(self, system, user, images, **kwargs):
            return region_response(user, '<script>alert("x")</script>'), TokenUsage()

    docs = describe_images({"Sheet": image_bytes()}, provider=HostileProvider())
    assert "<script>" not in docs["Sheet"]
    assert "&lt;script&gt;" in docs["Sheet"]


@pytest.mark.parametrize("invalid_kind", ["json", "observations", "characters"])
def test_contract_retry_regenerates_same_crop_and_preserves_both_attempts(invalid_kind):
    class RetryProvider(NativeProvider):
        def generate_with_images(self, system, user, images, **kwargs):
            self.calls.append((system, user, images))
            if len(self.calls) == 1:
                if invalid_kind == "json":
                    raw = "not JSON: rejected raw response"
                elif invalid_kind == "characters":
                    raw = region_response(user, "X" * 81)
                else:
                    reply = json.loads(region_response(user, "Rejected quotation"))
                    reply["regions"][0]["observations"] *= 4
                    raw = json.dumps(reply)
            else:
                raw = region_response(user, "Accepted quotation")
            self.vision_receipt = {"response_id": f"response-{len(self.calls)}"}
            return raw, TokenUsage(20, 10, 1)

    provider = RetryProvider()
    checks, usage = {}, TokenUsage()
    docs = describe_images(
        {"Sheet": image_bytes()},
        provider=provider,
        validation_results=checks,
        usage=usage,
        token_budget=60,
    )
    batch = checks["Sheet"]["pages"][0]["batches"][0]
    first, final = batch["attempts"]
    assert provider.calls[0][2] == provider.calls[1][2]
    assert len(provider.calls[0][2]) == BATCH_SIZE == 1
    retry_prompt = json.loads(provider.calls[1][1])
    assert retry_prompt["regions"] == json.loads(provider.calls[0][1])["regions"]
    assert retry_prompt["contract_retry"]["attempt"] == 2
    assert first["raw_response"] not in provider.calls[1][1]
    assert first["status"] == "invalid_response" and first["contract_error"]
    assert final["status"] == "complete"
    assert first["prompt_text_sha256"] != final["prompt_text_sha256"]
    assert first["json_prompt_sha256"] != final["json_prompt_sha256"]
    assert first["provider_receipt"]["response_id"] == "response-1"
    assert final["provider_receipt"]["response_id"] == "response-2"
    assert (
        batch["usage_total"]
        == vars(usage)
        == {
            "input_tokens": 40,
            "output_tokens": 20,
            "requests": 2,
            "estimated": False,
            "model": "",
            "provider": "",
        }
    )
    assert batch["final_attempt"] == 2 and batch["status"] == "complete"
    assert "contract_error" not in batch
    for key in (
        "user_prompt",
        "json_prompt_sha256",
        "prompt_text_sha256",
        "raw_response",
        "content_sha256",
        "usage",
        "provider_receipt",
    ):
        assert batch[key] == final[key]
    assert (
        first["content_sha256"]
        == hashlib.sha256(first["raw_response"].encode()).hexdigest()
    )
    if invalid_kind == "characters":
        assert (
            len(
                json.loads(first["raw_response"])["regions"][0]["observations"][0][
                    "text"
                ]
            )
            == 81
        )
    if invalid_kind == "observations":
        assert len(json.loads(first["raw_response"])["regions"][0]["observations"]) == 4
    assert "Accepted quotation" in docs["Sheet"]
    assert "Rejected quotation" not in docs["Sheet"]
    assert checks["Sheet"]["pages"][0]["completed_regions"] == ["r0-c0"]


@pytest.mark.parametrize("budget, expected_calls", [(30, 1), (59, 2)])
def test_invalid_attempt_cannot_retry_or_complete_over_budget(budget, expected_calls):
    class InvalidFirstProvider(NativeProvider):
        def generate_with_images(self, system, user, images, **kwargs):
            self.calls.append((system, user, images))
            return (
                "invalid JSON" if len(self.calls) == 1 else region_response(user)
            ), TokenUsage(20, 10, 1)

    provider = InvalidFirstProvider()
    checks, usage = {}, TokenUsage()
    with pytest.raises(AiDocError, match="budget") as caught:
        describe_images(
            {"Sheet": image_bytes()},
            provider=provider,
            validation_results=checks,
            usage=usage,
            token_budget=budget,
        )
    assert len(provider.calls) == expected_calls
    assert usage.total == caught.value.usage.total == expected_calls * 30
    batch = checks["Sheet"]["pages"][0]["batches"][0]
    assert batch["status"] == "failed" and batch["failure_code"] == "token_budget"
    assert len(batch["attempts"]) == expected_calls
    assert checks["Sheet"]["status"] == "incomplete"
    assert checks["Sheet"]["pages"][0]["completed_regions"] == []


@pytest.mark.parametrize(
    "reply", [" ", AiDocError("transport", usage=TokenUsage(20, 10, 1))]
)
def test_non_contract_failure_has_no_retry_and_keeps_usage(reply):
    class FailureProvider(NativeProvider):
        def generate_with_images(self, system, user, images, **kwargs):
            self.calls.append((system, user, images))
            if isinstance(reply, AiDocError):
                raise reply
            return reply, TokenUsage(20, 10, 1)

    provider = FailureProvider()
    checks, usage = {}, TokenUsage()
    with pytest.raises(AiDocError):
        describe_images(
            {"Sheet": image_bytes()},
            provider=provider,
            validation_results=checks,
            usage=usage,
        )
    assert len(provider.calls) == 1 and usage.total == 30
    batch = checks["Sheet"]["pages"][0]["batches"][0]
    assert len(batch["attempts"]) == 1 and batch["status"] == "failed"
    assert batch["usage_total"] == vars(usage)
    assert checks["Sheet"]["pages"][0]["completed_regions"] == []
