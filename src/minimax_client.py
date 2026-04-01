import os
import base64
import re
import time
import requests
from dotenv import load_dotenv

load_dotenv()

# MiniMax API Configuration
MINIMAX_API_KEY = "sk-cp-your-minimax-api-key-here"
MINIMAX_BASE_URL = "https://api.minimaxi.com"

_client = None


def _clean_thinking_content(content: str) -> str:
    """Remove thinking/reasoning content from MiniMax response.

    MiniMax-M2.7 reasoning model includes thinking process in responses.
    This function removes the thinking sections (marked with special tokens).
    """
    if not content:
        return content

    # Remove thinking content wrapped in special tags
    # Common patterns: <think>...</think>, Treasury:..., or similar
    patterns = [
        r'<think>.*?</think>',
        r' Treasury:.*?\n\n',
        r'思考过程:.*?\n\n',
        r'Thinking Process:.*?\n\n',
    ]

    cleaned = content
    for pattern in patterns:
        cleaned = re.sub(pattern, '', cleaned, flags=re.DOTALL | re.IGNORECASE)

    # Clean up leading whitespace/newlines
    cleaned = cleaned.strip()

    return cleaned if cleaned else content


class MiniMaxClient:
    """MiniMax API client wrapper."""

    def __init__(self, api_key: str, base_url: str):
        self.api_key = api_key
        self.base_url = base_url
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        }

    def chat_completion(
        self,
        model: str,
        messages: list[dict],
        temperature: float = 0.7,
    ) -> dict:
        """Call MiniMax chat completion API."""
        url = f"{self.base_url}/v1/chat/completions"
        payload = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
        }
        response = requests.post(url, headers=self.headers, json=payload)
        response.raise_for_status()
        return response.json()

    def image_generation(self, model: str, prompt: str) -> dict:
        """Call MiniMax image generation API."""
        url = f"{self.base_url}/v1/image_generation"
        payload = {
            "model": model,
            "prompt": prompt,
        }
        response = requests.post(url, headers=self.headers, json=payload)
        response.raise_for_status()
        return response.json()


def get_client() -> MiniMaxClient:
    """Get or create MiniMax client instance."""
    global _client
    if _client is None:
        api_key = os.getenv("MINIMAX_API_KEY", MINIMAX_API_KEY)
        if not api_key:
            raise ValueError("MINIMAX_API_KEY not found in environment variables")
        _client = MiniMaxClient(api_key=api_key, base_url=MINIMAX_BASE_URL)
    return _client


def generate_text(model: str, system_prompt: str, user_prompt: str, max_retries: int = 3) -> str:
    """Generate text using MiniMax API.

    Args:
        model: The model to use (e.g., 'minimax-text-01', 'abab6.5s-chat')
        system_prompt: System instruction for the model
        user_prompt: User's input prompt
        max_retries: Maximum number of retries for 529 errors

    Returns:
        Generated text response
    """
    client = get_client()
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    last_error = None
    for attempt in range(max_retries):
        try:
            response = client.chat_completion(
                model=model,
                messages=messages,
                temperature=0.7,
            )

            # Handle cases where choices might be null or missing
            choices = response.get("choices")
            if choices is None or len(choices) == 0:
                # Try to get reasoning_content from base_resp if available
                base_resp = response.get("base_resp", {})
                if base_resp.get("status_code", 0) != 0:
                    raise Exception(f"MiniMax API error: {base_resp.get('status_msg')}")
                # If reasoning_model returns reasoning_content
                reasoning = response.get("reasoning_content") or response.get("reasoning_content", "")
                if reasoning:
                    return _clean_thinking_content(reasoning)
                raise Exception(f"MiniMax API returned no response: {response}")

            raw_content = choices[0]["message"]["content"]
            return _clean_thinking_content(raw_content)

        except requests.exceptions.HTTPError as e:
            last_error = e
            if e.response is not None and e.response.status_code == 529:
                # 529 = overloaded, retry with exponential backoff
                wait_time = 2 ** attempt  # 2s, 4s, 8s
                print(f"[MiniMax] 529 过载错误，{wait_time}秒后重试 ({attempt + 1}/{max_retries})...")
                time.sleep(wait_time)
                continue
            else:
                raise
        except Exception as e:
            last_error = e
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt
                print(f"[MiniMax] 错误：{e}，{wait_time}秒后重试...")
                time.sleep(wait_time)
            else:
                raise

    # All retries exhausted
    raise Exception(f"MiniMax API 请求失败，已重试 {max_retries} 次：{last_error}")


def generate_text_with_images(
    model: str, system_prompt: str, user_prompt: str, image_paths: list[str]
) -> str:
    """Send text + images to MiniMax and get text response.

    Uses MiniMax's vision capability by passing images as base64.

    Args:
        model: The model to use for vision (should support image input)
        system_prompt: System instruction for the model
        user_prompt: User's input prompt
        image_paths: List of paths to image files

    Returns:
        Generated text response
    """
    client = get_client()

    # Build message content with text and images
    content = []

    # Add images as base64
    for img_path in image_paths:
        with open(img_path, "rb") as f:
            img_bytes = f.read()
        img_base64 = base64.b64encode(img_bytes).decode("utf-8")

        # Determine MIME type
        ext = img_path.lower().split(".")[-1]
        mime_type = "image/png" if ext == "png" else "image/jpeg"
        if ext == "gif":
            mime_type = "image/gif"
        elif ext == "webp":
            mime_type = "image/webp"

        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{mime_type};base64,{img_base64}"
            }
        })

    # Add text content
    content.append({
        "type": "text",
        "text": user_prompt
    })

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": content},
    ]

    response = client.chat_completion(
        model=model,
        messages=messages,
        temperature=0.3,
    )
    raw_content = response["choices"][0]["message"]["content"]
    return _clean_thinking_content(raw_content)


def generate_image(model: str, prompt: str) -> bytes | None:
    """Generate an image using MiniMax API.

    Args:
        model: The image model to use (e.g., 'image-01')
        prompt: Description of the image to generate

    Returns:
        Generated image as bytes, or None if generation failed
    """
    client = get_client()
    response = client.image_generation(model=model, prompt=prompt)

    # Check for API errors
    base_resp = response.get("base_resp", {})
    if base_resp.get("status_code", 0) != 0:
        print(f"MiniMax image generation error: {base_resp.get('status_msg')}")
        return None

    # MiniMax returns image URLs
    data = response.get("data")
    if data is not None and isinstance(data, dict) and "image_urls" in data:
        image_urls = data["image_urls"]
        if image_urls and len(image_urls) > 0:
            image_url = image_urls[0]
            img_response = requests.get(image_url)
            img_response.raise_for_status()
            return img_response.content

    return None


def generate_image_minimax(
    prompt: str,
    width: int = 1024,
    height: int = 1024,
    model: str = "image-01",
) -> bytes | None:
    """Generate an image using MiniMax API (wrapper for generate_image).

    Args:
        prompt: Description of the image to generate
        width: Image width (MiniMax API may not support custom sizes)
        height: Image height (MiniMax API may not support custom sizes)
        model: The image model to use

    Returns:
        Generated image as bytes, or None if generation failed
    """
    # MiniMax image API doesn't support custom dimensions directly
    # We use the default and let the API handle it
    return generate_image(model=model, prompt=prompt)