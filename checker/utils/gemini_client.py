"""Gemini API client with retry, logging, and cost tracking."""
import json
import threading
import time
from pathlib import Path

from google import genai
from google.genai import types as genai_types

from checker.config import (
    GEMINI_API_KEY, FLASH_MODEL, PRO_MODEL,
    FLASH_TEMPERATURE, PRO_TEMPERATURE,
    MAX_OUTPUT_TOKENS, MAX_RETRIES, RETRY_BASE_DELAY,
    COST_ESTIMATES,
)


class GeminiClient:
    """Wrapper around google-genai with retry, cost tracking, and JSON validation."""

    def __init__(self):
        self.client = genai.Client(api_key=GEMINI_API_KEY)
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.total_calls = 0
        self.call_log: list[dict] = []
        self._lock = threading.Lock()

    def call_flash(self, prompt: str, *, temperature: float | None = None) -> str:
        """Text-only call to Flash model."""
        return self._call(
            model=FLASH_MODEL,
            contents=prompt,
            temperature=temperature or FLASH_TEMPERATURE,
        )

    def call_flash_multimodal(self, parts: list, *, temperature: float | None = None) -> str:
        """Multimodal call to Flash model (text + images)."""
        return self._call(
            model=FLASH_MODEL,
            contents=parts,
            temperature=temperature or FLASH_TEMPERATURE,
        )

    def call_pro(self, prompt: str, *, temperature: float | None = None) -> str:
        """Text-only call to Pro model."""
        return self._call(
            model=PRO_MODEL,
            contents=prompt,
            temperature=temperature or PRO_TEMPERATURE,
        )

    def call_pro_multimodal(self, parts: list, *, temperature: float | None = None) -> str:
        """Multimodal call to Pro model (text + images)."""
        return self._call(
            model=PRO_MODEL,
            contents=parts,
            temperature=temperature or PRO_TEMPERATURE,
        )

    def _call(self, *, model: str, contents, temperature: float) -> str:
        """Core call with retry and logging."""
        for attempt in range(MAX_RETRIES):
            try:
                t0 = time.time()
                response = self.client.models.generate_content(
                    model=model,
                    contents=contents,
                    config={"temperature": temperature, "max_output_tokens": MAX_OUTPUT_TOKENS},
                )
                elapsed = time.time() - t0

                text = response.text or ""
                usage = getattr(response, "usage_metadata", None)
                in_tok = getattr(usage, "prompt_token_count", 0) or 0
                out_tok = getattr(usage, "candidates_token_count", 0) or 0

                with self._lock:
                    self.total_input_tokens += in_tok
                    self.total_output_tokens += out_tok
                    self.total_calls += 1
                    self.call_log.append({
                        "model": model,
                        "input_tokens": in_tok,
                        "output_tokens": out_tok,
                        "latency_s": round(elapsed, 1),
                    })

                return text

            except Exception as e:
                if attempt < MAX_RETRIES - 1:
                    wait = RETRY_BASE_DELAY ** (attempt + 1)
                    print(f"    WARN: Gemini error ({e}), retry in {wait}s...")
                    time.sleep(wait)
                else:
                    raise

        raise RuntimeError("Unreachable")

    def make_image_part(self, image_path: Path) -> genai_types.Part:
        """Create a Part object from an image file for multimodal calls."""
        data = image_path.read_bytes()
        mime = "image/png" if image_path.suffix.lower() == ".png" else "image/jpeg"
        return genai_types.Part.from_bytes(data=data, mime_type=mime)

    def estimated_cost(self) -> dict:
        """Return estimated cost breakdown."""
        cost_by_model: dict[str, float] = {}
        for entry in self.call_log:
            m = entry["model"]
            rates = COST_ESTIMATES.get(m, {"input": 0, "output": 0})
            cost = (entry["input_tokens"] * rates["input"] +
                    entry["output_tokens"] * rates["output"]) / 1_000_000
            cost_by_model[m] = cost_by_model.get(m, 0) + cost
        return {
            "total_calls": self.total_calls,
            "total_input_tokens": self.total_input_tokens,
            "total_output_tokens": self.total_output_tokens,
            "cost_by_model": cost_by_model,
            "total_cost_usd": sum(cost_by_model.values()),
        }

    def print_cost_summary(self):
        """Print cost summary to stdout."""
        info = self.estimated_cost()
        print(f"\n  API Usage:")
        print(f"    Calls: {info['total_calls']}")
        print(f"    Tokens: {info['total_input_tokens']:,} in / {info['total_output_tokens']:,} out")
        for m, c in info["cost_by_model"].items():
            print(f"    {m}: ~${c:.3f}")
        print(f"    Total: ~${info['total_cost_usd']:.3f}")
