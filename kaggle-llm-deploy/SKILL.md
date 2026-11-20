---
name: kaggle-llm-deploy
description: Deploy LLM models on Kaggle kernels with free GPU acceleration and expose via ngrok. Use when the user wants to run LLM inference on Kaggle, deploy GGUF models, or needs a temporary GPU API endpoint.
---

# Kaggle LLM Deploy

Deploy GGUF-format LLM models on Kaggle kernels using llama.cpp with GPU acceleration, exposed as HTTP API via ngrok.

## When to Use

- User wants to deploy an LLM model with free GPU
- User asks about Kaggle GPU inference
- User needs a temporary LLM API endpoint
- User mentions GGUF models or llama.cpp on Kaggle

## Core Workflow

1. Create Kaggle notebook with GPU accelerator (T4 x2 recommended)
2. Install llama.cpp using the automated installer
3. Download GGUF model from HuggingFace
4. Start llama.cpp server with GPU layers and ngrok tunnel
5. Test API with simple prompts

## Key Configuration

### Installation

Run this in a Kaggle notebook cell:

    curl -L https://github.com/Pluralise/llama.app/raw/main/install.sh | sh

This automatically detects GPU and builds with CUDA support. The build takes about 5-10 minutes.

### Server Parameters

    ./llama.cpp/llama-server \
        --model /kaggle/working/model.gguf \
        --host 0.0.0.0 \
        --port 7860 \
        --n-gpu-layers 99 \
        --ctx-size 32768 \
        --threads 4

Key flags:
- n-gpu-layers 99: Offload all layers to GPU (use -1 for CPU-only)
- ctx-size 32768: Context window (32k works well on T4)
- threads 4: CPU threads for non-GPU operations

### Ngrok Tunnel

    pip install pyngrok
    from pyngrok import ngrok
    ngrok.set_auth_token("YOUR_NGROK_TOKEN")
    public_url = ngrok.connect(7860).public_url
    print(f"API URL: {public_url}")

## Constraints

- 9-hour limit: Kaggle kernels auto-stop after 9 hours
- Model size: Must download within time limit. Large models (larger than 10GB) may timeout
- VRAM: T4 has 16GB per GPU. Qwen2.5-7B-Q4_K_M fits (~8GB)
- Context window: Larger context requires more VRAM. 200k needs more than 40GB VRAM

## Tested Models

| Model | Quantization | VRAM | Context | Speed |
|-------|-------------|------|---------|-------|
| Qwen3.5-0.8B | Default | ~1GB | 32k | 132-159 t/s |
| Qwen2.5-7B | Q4_K_M | ~8GB | 131k | 35-48 t/s |

## Common Issues

- CPU-only fallback: If GPU not detected, speed drops to about 10 t/s. Ensure llama.app/install.sh is used, not manual build.
- Context OOM: Reduce ctx-size if out of memory errors occur. Try 16384 or 8192.
- Model download timeout: Use smaller quantization or start early in kernel session. Prefer models under 5GB.

## API Usage

Once deployed, the API endpoint can be used with OpenAI-compatible format:

    import openai
    client = openai.OpenAI(
        base_url="YOUR_NGROK_URL/v1",
        api_key="kaggle"
    )
    response = client.chat.completions.create(
        model="local",
        messages=[{"role": "user", "content": "Hello!"}]
    )
    print(response.choices[0].message.content)
