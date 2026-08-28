import os
import logging
from openai import OpenAI

logger = logging.getLogger("officeai")

# OrcaRouter provides free access to DeepSeek, Qwen, and other models
# using the standard OpenAI API format.
_ORCAROUTER_BASE_URL = "https://api.orcarouter.ai/v1"

# Available free models:
#   deepseek/deepseek-v4-pro-free   ← best quality
#   deepseek/deepseek-v4-flash-free ← fastest
#   qwen/qwen3.8-27b-free


class LLMEngine:
    """
    OrcaRouter engine — free access to DeepSeek/Qwen via OpenAI-compatible API.
    Requires ORCAROUTER_API_KEY environment variable.
    """

    def __init__(self):
        api_key = os.getenv("ORCAROUTER_API_KEY")
        if not api_key:
            raise ValueError("ORCAROUTER_API_KEY environment variable is not set.")

        self.client = OpenAI(
            api_key=api_key,
            base_url=os.getenv("ORCAROUTER_BASE_URL", _ORCAROUTER_BASE_URL),
        )
        self.model = os.getenv("AI_MODEL", "deepseek/deepseek-v4-pro-free")
        self.max_tokens = int(os.getenv("OFFICEAI_MAX_TOKENS", "8192"))
        logger.info(f"OrcaRouter engine ready. Model: {self.model}")

    def stream_chat(self, messages: list):
        """Stream tokens from OrcaRouter API."""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=self.max_tokens,
            temperature=0.3,
            stream=True,
        )
        for chunk in response:
            delta = chunk.choices[0].delta
            if delta.content:
                yield delta.content
