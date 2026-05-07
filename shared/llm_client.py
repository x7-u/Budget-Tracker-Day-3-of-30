"""Anthropic client wrapper used across the 30 days.

- ask_claude / ask_claude_json: simple text + JSON helpers.
- Optional api_key override per-call (no persistent state).
- Cost estimate from the response.usage object.
- Retries on transient failures.
- Multimodal: prompts may be a string OR a list of content blocks (image + text).

Day 2 (Invoice Extractor) needs vision input. The image_block helper builds a
properly base64-encoded block; the call functions accept either a string or
a list-of-blocks for the user message content.
"""
from __future__ import annotations

import base64
import json
import logging
import re
import time
from dataclasses import dataclass

import anthropic

from shared.config import ANTHROPIC_API_KEY, CLAUDE_MODEL

log = logging.getLogger(__name__)

_default_client: anthropic.Anthropic | None = None

# USD per 1K tokens. Conservative estimates -- actuals vary by model.
# Updated for Claude 4.5/4.6/4.7 family.
PRICES: dict[str, tuple[float, float]] = {
    # model_substring : (input_per_1k, output_per_1k)
    "claude-haiku-4-5":  (0.001, 0.005),
    "claude-sonnet-4-6": (0.003, 0.015),
    "claude-sonnet-4-7": (0.003, 0.015),
    "claude-opus-4-7":   (0.015, 0.075),
}


@dataclass
class CallResult:
    text: str
    input_tokens: int
    output_tokens: int
    cache_read_tokens: int
    cache_creation_tokens: int
    cost_usd: float
    model: str


def _get_default_client() -> anthropic.Anthropic:
    global _default_client
    if _default_client is None:
        if not ANTHROPIC_API_KEY or ANTHROPIC_API_KEY.startswith("sk-ant-placeholder"):
            raise RuntimeError(
                "ANTHROPIC_API_KEY is missing or still a placeholder. "
                "Copy .env.example to .env and add your real key."
            )
        _default_client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    return _default_client


def _client_for(api_key: str | None) -> anthropic.Anthropic:
    if api_key:
        return anthropic.Anthropic(api_key=api_key)
    return _get_default_client()


def _price_for(model: str) -> tuple[float, float]:
    for key, prices in PRICES.items():
        if key in model:
            return prices
    return (0.003, 0.015)  # safe default


def _estimate_cost(usage, model: str) -> float:
    in_per_1k, out_per_1k = _price_for(model)
    inp = (getattr(usage, "input_tokens", 0) or 0)
    out = (getattr(usage, "output_tokens", 0) or 0)
    cache_read = (getattr(usage, "cache_read_input_tokens", 0) or 0)
    cache_create = (getattr(usage, "cache_creation_input_tokens", 0) or 0)
    # Cache reads cost ~10% of input. Cache creation costs ~125% of input.
    cost = (
        inp * in_per_1k / 1000
        + out * out_per_1k / 1000
        + cache_read * (in_per_1k * 0.10) / 1000
        + cache_create * (in_per_1k * 1.25) / 1000
    )
    return cost


def image_block(jpeg_or_png_bytes: bytes, media_type: str = "image/jpeg") -> dict:
    """Build a content block carrying a base64-encoded image.

    Use this together with text blocks to send multimodal prompts:
        prompt = [image_block(jpeg_bytes), {"type": "text", "text": "Extract..."}]
    """
    if media_type not in ("image/jpeg", "image/png", "image/gif", "image/webp"):
        raise ValueError(f"Unsupported media_type: {media_type}")
    return {
        "type": "image",
        "source": {
            "type": "base64",
            "media_type": media_type,
            "data": base64.standard_b64encode(jpeg_or_png_bytes).decode("ascii"),
        },
    }


def ask_claude_call(
    prompt: str | list[dict],
    *,
    system: str | list[dict] | None = None,
    max_tokens: int = 1000,
    model: str | None = None,
    api_key: str | None = None,
    retries: int = 2,
) -> CallResult:
    """Call Claude with the given prompt and return text + token/cost stats.

    `prompt` may be a string (text-only) or a list of content blocks
    (e.g. [image_block(...), {"type": "text", "text": "..."}]) for multimodal calls.
    """
    client = _client_for(api_key)
    model = model or CLAUDE_MODEL
    kwargs: dict = {
        "model": model,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": prompt}],  # str or list[dict]
    }
    if system is not None:
        kwargs["system"] = system

    last_exc: Exception | None = None
    for attempt in range(retries + 1):
        try:
            resp = client.messages.create(**kwargs)
            usage = getattr(resp, "usage", None)
            return CallResult(
                text=resp.content[0].text,
                input_tokens=getattr(usage, "input_tokens", 0) or 0,
                output_tokens=getattr(usage, "output_tokens", 0) or 0,
                cache_read_tokens=getattr(usage, "cache_read_input_tokens", 0) or 0,
                cache_creation_tokens=getattr(usage, "cache_creation_input_tokens", 0) or 0,
                cost_usd=_estimate_cost(usage, model),
                model=model,
            )
        except (anthropic.APIConnectionError, anthropic.APIStatusError) as e:
            last_exc = e
            if attempt < retries:
                time.sleep(0.5 * (2 ** attempt))
                continue
            raise
        except Exception as e:
            last_exc = e
            raise
    raise last_exc  # type: ignore[misc]


def ask_claude(prompt: str, system: str | None = None, max_tokens: int = 1000, **kwargs) -> str:
    """Backwards-compatible: text-only response."""
    return ask_claude_call(prompt, system=system, max_tokens=max_tokens, **kwargs).text


def ask_claude_json(
    prompt: str | list[dict],
    *,
    system: str | list[dict] | None = None,
    max_tokens: int = 1500,
    model: str | None = None,
    api_key: str | None = None,
) -> dict:
    """Ask for a JSON object. Tolerates ``` fences and stray prose. Returns the parsed dict."""
    sys_payload = system
    if isinstance(system, str):
        sys_payload = system + "\n\nReturn ONLY a single valid JSON object. No prose, no markdown fences."
    elif isinstance(system, list):
        # Append the JSON instruction to the last text block.
        sys_payload = list(system)
        if sys_payload and sys_payload[-1].get("type") == "text":
            sys_payload[-1] = {
                **sys_payload[-1],
                "text": sys_payload[-1]["text"] + "\n\nReturn ONLY a single valid JSON object. No prose, no markdown fences.",
            }
    res = ask_claude_call(prompt, system=sys_payload, max_tokens=max_tokens, model=model, api_key=api_key)
    return _parse_json_loose(res.text)


def ask_claude_json_with_stats(
    prompt: str | list[dict],
    *,
    system: str | list[dict] | None = None,
    max_tokens: int = 1500,
    model: str | None = None,
    api_key: str | None = None,
) -> tuple[dict, CallResult]:
    """Same as ask_claude_json but also returns the raw CallResult (cost stats)."""
    sys_payload = system
    if isinstance(system, str):
        sys_payload = system + "\n\nReturn ONLY a single valid JSON object. No prose, no markdown fences."
    elif isinstance(system, list):
        sys_payload = list(system)
        if sys_payload and sys_payload[-1].get("type") == "text":
            sys_payload[-1] = {
                **sys_payload[-1],
                "text": sys_payload[-1]["text"] + "\n\nReturn ONLY a single valid JSON object. No prose, no markdown fences.",
            }
    res = ask_claude_call(prompt, system=sys_payload, max_tokens=max_tokens, model=model, api_key=api_key)
    return _parse_json_loose(res.text), res


def _parse_json_loose(raw: str) -> dict:
    cleaned = raw.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*|\s*```$", "", cleaned, flags=re.MULTILINE).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        m = re.search(r"\{.*\}", cleaned, flags=re.DOTALL)
        if m:
            return json.loads(m.group(0))
        raise
