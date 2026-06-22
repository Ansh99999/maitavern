# 07 — Providers & Models

Connections are grouped by **purpose**: **Text · Image · Voice**. A text connection is created as either a **Provider** (simple) or a **Router** (advanced).

## Text → Provider (simple direct connection)

- **Name**
- **Method** — sets request/response/stream shape + auth header style:
  - Anthropic style → `v1/messages`
  - OpenAI-compatible → `v1/chat/completions`
  - OpenAI Responses → `v1/responses`
  - Gemini style → `generateContent`
- **Base URL**
- **API key(s)** — `＋ Add key`; with more than one key, a **rotation strategy** appears: round robin · random · failover-on-error · weighted.
- **Test connection** — validates and auto-lists available models.

Keys are encrypted at rest, masked, never logged.

## Text → Router (superset of Provider)

For aggregators/gateways (OpenRouter, LiteLLM) and power routing — all Provider fields plus:

- **Per-model pricing** (auto-fetch where exposed, else manual)
- **Cost & token tracking** → feeds the Usage dashboard
- **Fallback chain** (on error/timeout/rate-limit)
- **Custom parameters** (merged into the request body)
- **Param include/exclude** — the compatibility valve: a preset defines sampler values, but the connection decides which actually get sent (some endpoints reject `top_k`/penalties). The Log viewer shows the final filtered payload.

## Image models

Add an image provider — method (OpenAI Images · Stable Diffusion/A1111 · ComfyUI · Gemini image · custom), base URL, key, model, default size/steps/sampler. Powers character art, scene images, and the Gallery.

## Voice models

Add a voice provider — **TTS** (OpenAI TTS · ElevenLabs · custom) and **STT** (Whisper · custom) for read-aloud and voice input.

## Model Roles

One mapping assigns a **connection + model to each job** — Roleplay (default) · **Mai assistant** (dedicated) · **Memory agent** (user-configurable) · Embeddings (RAG) · Image · Voice (TTS/STT).

## Usage & Costs dashboard

Spend + tokens per chat / character / model / day, with optional budget alerts; fed by pricing + token counts.

> Other productive bits: Test Connection auto-fetches the model list; per-key health (rate-limited keys skipped by rotation); smart default include/exclude per method.

## Confirmed decisions

- **Model Roles = both** — a central Model-Roles map is the source of truth, with inline shortcuts at each feature pointing back to it.
- **Cost tracking = plain Providers may add optional manual per-model pricing** to unlock the same cost tracking as Router. Token counts always shown regardless of type.

## Network settings

Streaming method (native plugin vs fetch), allow cleartext/LAN, Anthropic browser header, request timeout, proxy. See risks in [00-vision-and-architecture.md](00-vision-and-architecture.md).
