"""
Aliyun Vision Client for text generation with images.
Uses OpenAI-compatible API format.
"""

import base64
from pathlib import Path
from typing import List, Optional

import requests


# Aliyun API Configuration
ALIYUN_API_KEY = "sk-your-aliyun-api-key-here"
ALIYUN_BASE_URL = "https://coding.dashscope.aliyuncs.com"
DEFAULT_MODEL = "qwen3.5-plus"


def encode_image_to_base64(image_path: str) -> str:
    """
    Encode an image file to base64 string.

    Args:
        image_path: Path to the image file

    Returns:
        Base64 encoded string with data URI prefix
    """
    path = Path(image_path)
    if not path.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    # Determine MIME type based on file extension
    suffix = path.suffix.lower()
    mime_types = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.gif': 'image/gif',
        '.webp': 'image/webp',
    }
    mime_type = mime_types.get(suffix, 'image/jpeg')

    # Read and encode the image
    with open(path, 'rb') as f:
        image_data = f.read()

    base64_data = base64.b64encode(image_data).decode('utf-8')
    return f"data:{mime_type};base64,{base64_data}"


def generate_text_with_images(
    model: str,
    system_prompt: str,
    user_prompt: str,
    image_paths: Optional[List[str]] = None
) -> str:
    """
    Generate text with vision capabilities using Aliyun API.

    Args:
        model: The model to use (e.g., "qwen3.5-plus")
        system_prompt: System prompt for the conversation
        user_prompt: User prompt/message
        image_paths: Optional list of image file paths to include

    Returns:
        Generated text response from the model
    """
    url = f"{ALIYUN_BASE_URL}/v1/chat/completions"

    headers = {
        "Authorization": f"Bearer {ALIYUN_API_KEY}",
        "Content-Type": "application/json"
    }

    # Build the user message content
    user_content = []

    # Add images if provided
    if image_paths:
        for image_path in image_paths:
            base64_image = encode_image_to_base64(image_path)
            user_content.append({
                "type": "image_url",
                "image_url": {
                    "url": base64_image
                }
            })

    # Add text prompt
    user_content.append({
        "type": "text",
        "text": user_prompt
    })

    # Build the request body
    payload = {
        "model": model,
        "messages": [
            {
                "role": "system",
                "content": system_prompt
            },
            {
                "role": "user",
                "content": user_content
            }
        ]
    }

    # Make the API request
    response = requests.post(url, headers=headers, json=payload)
    response.raise_for_status()

    result = response.json()
    return result["choices"][0]["message"]["content"]


def generate_text(
    model: str,
    system_prompt: str,
    user_prompt: str
) -> str:
    """
    Generate text without images using Aliyun API.

    Args:
        model: The model to use (e.g., "qwen3.5-plus")
        system_prompt: System prompt for the conversation
        user_prompt: User prompt/message

    Returns:
        Generated text response from the model
    """
    return generate_text_with_images(model, system_prompt, user_prompt, image_paths=None)


if __name__ == "__main__":
    # Example usage
    model = DEFAULT_MODEL
    system_prompt = "You are a helpful assistant."
    user_prompt = "Hello, how are you?"

    result = generate_text(model, system_prompt, user_prompt)
    print(f"Response: {result}")