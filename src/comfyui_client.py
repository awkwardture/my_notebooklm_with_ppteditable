"""
ComfyUI client for local image generation.
Supports Chinese text rendering with Flux/SDXL models.
"""

import io
import json
import os
import time
import uuid
import random
import requests
from typing import Optional
from dotenv import load_dotenv

load_dotenv()

# ComfyUI Configuration
COMFYUI_URL = os.getenv("COMFYUI_URL", "http://127.0.0.1:8188")

# Z-Image-Turbo 模型配置 (最快)
Z_IMAGE_TURBO = "z_image_turbo_bf16.safetensors"

# Qwen Image 2512 模型配置 (文本生成图片，支持中文)
QWEN_UNET = "qwen_image_2512_fp8_e4m3fn.safetensors"
QWEN_CLIP = "qwen_2.5_vl_7b_fp8_scaled.safetensors"
QWEN_VAE = "qwen_image_vae.safetensors"
QWEN_LORA = "Qwen-Image-Lightning-4steps-V1.0.safetensors"
# Qwen CLIP 类型 - 如果 ComfyUI 不支持 qwen_image 类型，尝试其他类型
QWEN_CLIP_TYPE = "qwen_image"  # 或尝试 "sd3", "flux", 或留空

# Flux 模型配置 (高质量但慢)
FLUX_UNET = "flux1-fill-dev.safetensors"
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


def create_qwen_image_2512_workflow(
    prompt: str,
    negative_prompt: str = "低分辨率，低画质，肢体畸形，手指缺失，模糊，低质量，丑陋，畸形",
    width: int = 1920,
    height: int = 1080,
    steps: int = 50,
    cfg: float = 4.0,
    use_lora: bool = False,
    seed: int = None,
) -> dict:
    """Create a Qwen Image 2512 workflow for text-to-image generation.

    Based on: /data/gitrepo/ComfyUI/user/default/workflows/image_qwen_Image_2512.json

    Args:
        prompt: Text prompt
        negative_prompt: Negative prompt (Chinese by default)
        width: Image width
        height: Image height
        steps: Sampling steps (50 for standard, 4 for LoRA version)
        cfg: CFG scale (4.0 for standard, 1.0 for LoRA)
        use_lora: Use LoRA version for faster generation (4 steps)
        seed: Random seed
    """
    if seed is None:
        seed = random.randint(0, 2**32 - 1)

    # 基础工作流节点（按顺序构建，避免 ID 冲突）
    workflow = {}

    # UNET Loader - 加载 Qwen Image 2512 模型
    workflow["1"] = {
        "class_type": "UNETLoader",
        "inputs": {
            "unet_name": QWEN_UNET,
            "weight_dtype": "default"
        }
    }

    # CLIP Loader - 加载 Qwen CLIP
    # 注意：需要 ComfyUI 支持 qwen_image 类型
    # 如果报错，请确保你的 ComfyUI 是最新版本，或尝试修改 type 为 "sd3" 或 "flux"
    workflow["2"] = {
        "class_type": "CLIPLoader",
        "inputs": {
            "clip_name": QWEN_CLIP,
            "type": QWEN_CLIP_TYPE
        }
    }

    # VAE Loader
    workflow["3"] = {
        "class_type": "VAELoader",
        "inputs": {
            "vae_name": QWEN_VAE
        }
    }

    # Empty SD3 Latent Image (必须在 KSampler 之前)
    workflow["4"] = {
        "class_type": "EmptySD3LatentImage",
        "inputs": {
            "width": width,
            "height": height,
            "batch_size": 1
        }
    }

    # Positive prompt
    workflow["5"] = {
        "class_type": "CLIPTextEncode",
        "inputs": {
            "clip": ["2", 0],
            "text": prompt
        }
    }

    # Negative prompt
    workflow["6"] = {
        "class_type": "CLIPTextEncode",
        "inputs": {
            "clip": ["2", 0],
            "text": negative_prompt
        }
    }

    # 如果使用 LoRA，添加 LoraLoaderModelOnly
    if use_lora:
        # LoraLoaderModelOnly - 加载 LoRA
        workflow["7"] = {
            "class_type": "LoraLoaderModelOnly",
            "inputs": {
                "model": ["1", 0],
                "lora_name": QWEN_LORA,
                "strength_model": 1.0
            }
        }
        # KSampler - 4 steps for LoRA (cfg=1.0)
        workflow["8"] = {
            "class_type": "KSampler",
            "inputs": {
                "model": ["7", 0],  # 从 LoraLoaderModelOnly 输出
                "positive": ["5", 0],
                "negative": ["6", 0],
                "latent_image": ["4", 0],
                "seed": seed,
                "steps": 4,
                "cfg": 1.0,
                "sampler_name": "euler",
                "scheduler": "simple",
                "denoise": 1.0
            }
        }
    else:
        # ModelSamplingAuraFlow - Apply AuraFlow sampling with shift=3.1
        workflow["7"] = {
            "class_type": "ModelSamplingAuraFlow",
            "inputs": {
                "model": ["1", 0],
                "shift": 3.1
            }
        }
        # KSampler - standard 50 steps
        workflow["8"] = {
            "class_type": "KSampler",
            "inputs": {
                "model": ["7", 0],  # 从 ModelSamplingAuraFlow 输出
                "positive": ["5", 0],
                "negative": ["6", 0],
                "latent_image": ["4", 0],
                "seed": seed,
                "steps": steps,
                "cfg": cfg,
                "sampler_name": "euler",
                "scheduler": "simple",
                "denoise": 1.0
            }
        }

    # VAE Decode
    workflow["9"] = {
        "class_type": "VAEDecode",
        "inputs": {
            "samples": ["8", 0],
            "vae": ["3", 0]
        }
    }

    # Save Image
    workflow["10"] = {
        "class_type": "SaveImage",
        "inputs": {
            "images": ["9", 0],
            "filename_prefix": "PPT"
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
    steps: int = 50,
    seed: int = None,
    use_z_image_turbo: bool = True,
    use_qwen_2512: bool = False,
    use_qwen_fast: bool = False,
    use_flux: bool = False,
    model: str = None,
) -> bytes | None:
    """Generate an image using local ComfyUI service.

    Args:
        prompt: Text prompt for image generation
        negative_prompt: Negative prompt (used for Qwen Image 2512)
        width: Image width (default 1920 for 16:9)
        height: Image height (default 1080 for 16:9)
        steps: Sampling steps (Z-Image-Turbo recommends 20, Qwen 2512 recommends 50)
        seed: Random seed (None for random)
        use_z_image_turbo: Use Z-Image-Turbo model (fastest, best Chinese)
        use_qwen_2512: Use Qwen Image 2512 model (best Chinese support, 50 steps)
        use_qwen_fast: Use Qwen Image 2512 LoRA version (4 steps, faster)
        use_flux: Use Flux model (high quality but slow)
        model: Model checkpoint name (for SDXL)

    Returns:
        Generated image as bytes, or None if failed
    """
    try:
        # Create workflow
        print(f"[ComfyUI] 使用模型：use_z_image_turbo={use_z_image_turbo}, use_qwen_2512={use_qwen_2512}, use_qwen_fast={use_qwen_fast}, use_flux={use_flux}")
        if use_z_image_turbo:
            workflow = create_z_image_turbo_workflow(
                prompt=prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
            print(f"[ComfyUI] Z-Image-Turbo workflow created")
        elif use_qwen_2512:
            workflow = create_qwen_image_2512_workflow(
                prompt=prompt,
                negative_prompt=negative_prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
            print(f"[ComfyUI] Qwen Image 2512 workflow created (50 steps)")
        elif use_qwen_fast:
            workflow = create_qwen_image_2512_workflow(
                prompt=prompt,
                negative_prompt=negative_prompt,
                width=width,
                height=height,
                steps=4,
                seed=seed,
                use_lora=True,
            )
            print(f"[ComfyUI] Qwen Image 2512 LoRA workflow created (4 steps)")
        elif use_flux:
            workflow = create_flux_workflow(
                prompt=prompt,
                width=width,
                height=height,
                steps=steps,
                seed=seed,
            )
            print(f"[ComfyUI] Flux workflow created, UNET={FLUX_UNET}")
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
            print(f"[ComfyUI] SDXL workflow created, model={model}")

        # Queue prompt
        print(f"[ComfyUI] 提交工作流...")
        prompt_id, client_id = _queue_prompt(workflow)
        print(f"[ComfyUI] Prompt ID: {prompt_id}")

        # Wait for completion
        print(f"[ComfyUI] 等待完成...")
        history = _wait_for_completion(prompt_id, client_id)
        print(f"[ComfyUI] 完成，history keys: {history.keys()}")

        # Check for errors in status
        status = history.get("status", {})
        if status:
            print(f"[ComfyUI] Status: {status}")

        # Get output images
        outputs = history.get("outputs", {})
        print(f"[ComfyUI] outputs: {outputs.keys() if outputs else 'empty'}")
        for node_id, output in outputs.items():
            if "images" in output:
                for image_info in output["images"]:
                    filename = image_info["filename"]
                    subfolder = image_info.get("subfolder", "")
                    print(f"[ComfyUI] 获取图片：{filename}")
                    return _get_image(filename, subfolder)

        print("[ComfyUI] 错误：outputs 中没有 images")
        print(f"Full history: {json.dumps(history, indent=2)[:3000]}")
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