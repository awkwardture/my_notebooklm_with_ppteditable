"""
ComfyUI client for local image generation.
Supports Chinese text rendering with Flux/SDXL models.
"""

import io
import json
import time
import uuid
import random
import requests
from typing import Optional

# ComfyUI Configuration
COMFYUI_URL = "http://127.0.0.1:8188"

# Z-Image-Turbo 模型配置 (最快)
Z_IMAGE_TURBO = "z_image_turbo_bf16.safetensors"

# Qwen Image Edit 模型配置 (快速且支持中文)
QWEN_UNET = "qwen_image_edit_2509_fp8_e4m3fn.safetensors"
QWEN_CLIP = "qwen_2.5_vl_7b_fp8_scaled.safetensors"
QWEN_VAE = "qwen_image_vae.safetensors"

# Flux 模型配置 (高质量但慢)
FLUX_UNET = "flux1-krea-dev_fp8_scaled.safetensors"
FLUX_CLIP = "t5xxl_fp16.safetensors"
FLUX_VAE = "ae.safetensors"


def _queue_prompt(workflow: dict) -> tuple:
    """Send workflow to ComfyUI and return prompt_id."""
    client_id = str(uuid.uuid4())
    payload = {"prompt": workflow, "client_id": client_id}
    response = requests.post(f"{COMFYUI_URL}/prompt", json=payload)
    response.raise_for_status()
    return response.json()["prompt_id"], client_id


def _wait_for_completion(prompt_id: str, client_id: str, timeout: int = 600) -> dict:
    """Wait for workflow to complete and return history."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        response = requests.get(f"{COMFYUI_URL}/history/{prompt_id}")
        if response.status_code == 200:
            history = response.json()
            if prompt_id in history:
                return history[prompt_id]
        time.sleep(1)
    raise TimeoutError(f"ComfyUI workflow timed out after {timeout}s")


def _get_image(filename: str, subfolder: str = "", folder_type: str = "output") -> bytes:
    """Download generated image from ComfyUI."""
    params = {"filename": filename, "subfolder": subfolder, "type": folder_type}
    response = requests.get(f"{COMFYUI_URL}/view", params=params)
    response.raise_for_status()
    return response.content


def create_flux_workflow(
    prompt: str,
    negative_prompt: str = "",
    width: int = 1920,
    height: int = 1080,
    steps: int = 20,
    cfg: float = 3.5,
    seed: int = None,
) -> dict:
    """Create a Flux workflow for text-to-image generation."""
    if seed is None:
        seed = random.randint(0, 2**32 - 1)

    workflow = {
        # UNET Loader - 加载 Flux 模型
        "1": {
            "class_type": "UNETLoader",
            "inputs": {
                "unet_name": FLUX_UNET,
                "weight_dtype": "default"
            }
        },
        # CLIP Loader - 加载 T5 文本编码器
        "2": {
            "class_type": "CLIPLoader",
            "inputs": {
                "clip_name": FLUX_CLIP,
                "type": "flux2"
            }
        },
        # VAE Loader
        "3": {
            "class_type": "VAELoader",
            "inputs": {
                "vae_name": FLUX_VAE
            }
        },
        # Positive prompt
        "4": {
            "class_type": "CLIPTextEncode",
            "inputs": {
                "clip": ["2", 0],
                "text": prompt
            }
        },
        # Empty Latent
        "5": {
            "class_type": "EmptyLatentImage",
            "inputs": {
                "width": width,
                "height": height,
                "batch_size": 1
            }
        },
        # KSampler
        "6": {
            "class_type": "KSampler",
            "inputs": {
                "model": ["1", 0],
                "positive": ["4", 0],
                "negative": ["4", 0],  # Flux 不使用 negative prompt
                "latent_image": ["5", 0],
                "seed": seed,
                "steps": steps,
                "cfg": cfg,
                "sampler_name": "euler",
                "scheduler": "simple",
                "denoise": 1.0
            }
        },
        # VAE Decode
        "7": {
            "class_type": "VAEDecode",
            "inputs": {
                "samples": ["6", 0],
                "vae": ["3", 0]
            }
        },
        # Save Image
        "8": {
            "class_type": "SaveImage",
            "inputs": {
                "images": ["7", 0],
                "filename_prefix": "PPT"
            }
        }
    }
    return workflow


def create_qwen_image_workflow(
    prompt: str,
    width: int = 1920,
    height: int = 1080,
    steps: int = 15,
    cfg: float = 4.0,
    seed: int = None,
) -> dict:
    """Create a Qwen Image workflow for text-to-image generation."""
    if seed is None:
        seed = random.randint(0, 2**32 - 1)

    workflow = {
        # UNET Loader - 加载 Qwen Image 模型
        "1": {
            "class_type": "UNETLoader",
            "inputs": {
                "unet_name": QWEN_UNET,
                "weight_dtype": "fp8_e4m3fn"
            }
        },
        # CLIP Loader - 加载 Qwen CLIP
        "2": {
            "class_type": "CLIPLoader",
            "inputs": {
                "clip_name": QWEN_CLIP,
                "type": "qwen_image"
            }
        },
        # VAE Loader
        "3": {
            "class_type": "VAELoader",
            "inputs": {
                "vae_name": QWEN_VAE
            }
        },
        # Positive prompt
        "4": {
            "class_type": "CLIPTextEncode",
            "inputs": {
                "clip": ["2", 0],
                "text": prompt
            }
        },
        # Empty Latent
        "5": {
            "class_type": "EmptyLatentImage",
            "inputs": {
                "width": width,
                "height": height,
                "batch_size": 1
            }
        },
        # KSampler
        "6": {
            "class_type": "KSampler",
            "inputs": {
                "model": ["1", 0],
                "positive": ["4", 0],
                "negative": ["4", 0],
                "latent_image": ["5", 0],
                "seed": seed,
                "steps": steps,
                "cfg": cfg,
                "sampler_name": "euler",
                "scheduler": "simple",
                "denoise": 1.0
            }
        },
        # VAE Decode
        "7": {
            "class_type": "VAEDecode",
            "inputs": {
                "samples": ["6", 0],
                "vae": ["3", 0]
            }
        },
        # Save Image
        "8": {
            "class_type": "SaveImage",
            "inputs": {
                "images": ["7", 0],
                "filename_prefix": "PPT"
            }
        }
    }
    return workflow


def create_z_image_turbo_workflow(
    prompt: str,
    width: int = 1920,
    height: int = 1080,
    steps: int = 20,  # Z-Image-Turbo 推荐 20 步
    cfg: float = 1.0,  # cfg=1 for turbo models
    seed: int = None,
) -> dict:
    """Create a Z-Image-Turbo workflow for fast text-to-image generation.

    Based on official ComfyUI workflow template:
    - CLIPLoader: qwen_3_4b.safetensors (lumina2 type)
    - VAELoader: ae.safetensors
    - UNETLoader: z_image_turbo_bf16.safetensors
    - ModelSamplingAuraFlow with shift=3
    - ConditioningZeroOut for negative conditioning
    - EmptySD3LatentImage
    - KSampler: res_multistep sampler, simple scheduler
    """
    if seed is None:
        seed = random.randint(0, 2**32 - 1)

    workflow = {
        # UNET Loader - 加载 Z-Image-Turbo
        "1": {
            "class_type": "UNETLoader",
            "inputs": {
                "unet_name": Z_IMAGE_TURBO,
                "weight_dtype": "default"
            }
        },
        # CLIP Loader - lumina2 类型
        "2": {
            "class_type": "CLIPLoader",
            "inputs": {
                "clip_name": "qwen_3_4b.safetensors",
                "type": "lumina2",
                "device": "default"
            }
        },
        # VAE Loader - ae.safetensors
        "3": {
            "class_type": "VAELoader",
            "inputs": {
                "vae_name": "ae.safetensors"
            }
        },
        # CLIP Text Encode - Positive prompt
        "4": {
            "class_type": "CLIPTextEncode",
            "inputs": {
                "clip": ["2", 0],
                "text": prompt
            }
        },
        # ConditioningZeroOut - Negative conditioning (required for turbo)
        "5": {
            "class_type": "ConditioningZeroOut",
            "inputs": {
                "conditioning": ["4", 0]
            }
        },
        # ModelSamplingAuraFlow - Apply AuraFlow sampling with shift=3
        "6": {
            "class_type": "ModelSamplingAuraFlow",
            "inputs": {
                "model": ["1", 0],
                "shift": 3
            }
        },
        # Empty SD3 Latent Image
        "7": {
            "class_type": "EmptySD3LatentImage",
            "inputs": {
                "width": width,
                "height": height,
                "batch_size": 1
            }
        },
        # KSampler - res_multistep sampler
        "8": {
            "class_type": "KSampler",
            "inputs": {
                "model": ["6", 0],  # From ModelSamplingAuraFlow
                "positive": ["4", 0],  # From CLIPTextEncode
                "negative": ["5", 0],  # From ConditioningZeroOut
                "latent_image": ["7", 0],
                "seed": seed,
                "steps": steps,
                "cfg": cfg,
                "sampler_name": "res_multistep",
                "scheduler": "simple",
                "denoise": 1.0
            }
        },
        # VAE Decode
        "9": {
            "class_type": "VAEDecode",
            "inputs": {
                "samples": ["8", 0],
                "vae": ["3", 0]
            }
        },
        # Save Image
        "10": {
            "class_type": "SaveImage",
            "inputs": {
                "images": ["9", 0],
                "filename_prefix": "PPT"
            }
        }
    }
    return workflow


def create_sdxl_workflow(
    prompt: str,
    negative_prompt: str = "",
    width: int = 1920,
    height: int = 1080,
    steps: int = 20,
    cfg: float = 7.0,
    seed: int = None,
    model: str = "sd_xl_base_1.0.safetensors",
) -> dict:
    """Create an SDXL workflow for text-to-image generation."""
    if seed is None:
        seed = random.randint(0, 2**32 - 1)

    workflow = {
        "1": {
            "class_type": "CheckpointLoaderSimple",
            "inputs": {
                "ckpt_name": model
            }
        },
        "2": {
            "class_type": "CLIPTextEncode",
            "inputs": {
                "clip": ["1", 1],
                "text": prompt
            }
        },
        "3": {
            "class_type": "CLIPTextEncode",
            "inputs": {
                "clip": ["1", 1],
                "text": negative_prompt or "low quality, worst quality, blurry"
            }
        },
        "4": {
            "class_type": "EmptyLatentImage",
            "inputs": {
                "width": width,
                "height": height,
                "batch_size": 1
            }
        },
        "5": {
            "class_type": "KSampler",
            "inputs": {
                "model": ["1", 0],
                "positive": ["2", 0],
                "negative": ["3", 0],
                "latent_image": ["4", 0],
                "seed": seed,
                "steps": steps,
                "cfg": cfg,
                "sampler_name": "euler",
                "scheduler": "normal",
                "denoise": 1.0
            }
        },
        "6": {
            "class_type": "VAEDecode",
            "inputs": {
                "samples": ["5", 0],
                "vae": ["1", 2]
            }
        },
        "7": {
            "class_type": "SaveImage",
            "inputs": {
                "images": ["6", 0],
                "filename_prefix": "PPT"
            }
        }
    }
    return workflow


def generate_image_comfyui(
    prompt: str,
    negative_prompt: str = "",
    width: int = 1920,
    height: int = 1080,
    steps: int = 20,
    seed: int = None,
    use_z_image_turbo: bool = True,
    use_qwen: bool = False,
    use_flux: bool = False,
    model: str = None,
) -> bytes | None:
    """Generate an image using local ComfyUI service.

    Args:
        prompt: Text prompt for image generation
        negative_prompt: Negative prompt (not used for Flux/Qwen/Turbo)
        width: Image width (default 1920 for 16:9)
        height: Image height (default 1080 for 16:9)
        steps: Sampling steps (Z-Image-Turbo recommends 20)
        seed: Random seed (None for random)
        use_z_image_turbo: Use Z-Image-Turbo model (fastest, best Chinese)
        use_qwen: Use Qwen Image model (good Chinese support)
        use_flux: Use Flux model (high quality but slow)
        model: Model checkpoint name (for SDXL)

    Returns:
        Generated image as bytes, or None if failed
    """
    try:
        # Create workflow
        if use_z_image_turbo:
            workflow = create_z_image_turbo_workflow(
                prompt=prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
        elif use_qwen:
            workflow = create_qwen_image_workflow(
                prompt=prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
        elif use_flux:
            workflow = create_flux_workflow(
                prompt=prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
        else:
            workflow = create_sdxl_workflow(
                prompt=prompt,
                negative_prompt=negative_prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
                model=model or "sd_xl_base_1.0.safetensors",
            )

        # Queue prompt
        prompt_id, client_id = _queue_prompt(workflow)

        # Wait for completion
        history = _wait_for_completion(prompt_id, client_id)

        # Get output images
        outputs = history.get("outputs", {})
        for node_id, output in outputs.items():
            if "images" in output:
                for image_info in output["images"]:
                    filename = image_info["filename"]
                    subfolder = image_info.get("subfolder", "")
                    return _get_image(filename, subfolder)

        return None

    except Exception as e:
        print(f"ComfyUI generation error: {e}")
        return None


if __name__ == "__main__":
    # Test
    print("Testing ComfyUI client...")
    img = generate_image_comfyui(
        prompt="professional presentation slide background, clean modern design, blue theme, geometric shapes, no text",
        width=1920,
        height=1080,
        steps=15,
    )
    if img:
        with open("/tmp/comfyui_test.png", "wb") as f:
            f.write(img)
        print(f"Image saved to /tmp/comfyui_test.png, size: {len(img)} bytes")
    else:
        print("Failed to generate image")