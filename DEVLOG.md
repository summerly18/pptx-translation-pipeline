# Development Log — PPTX Translation Pipeline (FR→KO)

A local LLM pipeline that extracts text from PowerPoint decks, translates FR→KO
with a glossary, exports a review spreadsheet, and re-inserts the approved
translation back into the original PPTX.

Stack: Python, python-pptx, openpyxl, llama.cpp (Qwen3.5-9B-Q4_K_M), RTX 5070 Laptop 8GB.

---

## Day 1 — 2026-05-29

### Summary
Built the full pipeline end-to-end and validated it on live client decks
(Archives du Maroc action plans). Restructured into discrete, resumable stages.
Settled model and prompting strategy after several inference-engine issues.

### Pipeline architecture (final)
Four stages, each writing to its own folder so work survives interruption:

```
input/         원본 PPTX
1_extracted/   추출된 텍스트 (번역 전)   ← 1_extract.py
2_translated/  Qwen 초벌 번역            ← 2_translate.py
3_reviewed/    감수 완료 (Claude+human)
4_final/       번역 삽입된 PPTX          ← reinsert.py
```

Rationale: extraction is fast and reliable, translation is slow and fragile.
Separating them means a translation crash never loses extracted data, and the
review stage has a stable artifact to work from.

### Issues & Resolutions

| # | Issue | Root Cause | Resolution |
|---|-------|-----------|------------|
| 1 | Mistral Nemo 12B: missing translations, English output, prompt leakage | Weak FR→KO coverage; strong only on FR→EN | Switched to Qwen3.5 9B |
| 2 | Qwen3.5 load failure: `missing tensor 'blk.32.ssm_conv1d.weight'` | SSM/hybrid architecture unsupported by installed llama.cpp | Updated llama.cpp to build b9374 |
| 3 | Empty `content`, output only in `reasoning_content` | Qwen3.5 hybrid reasoning model; thinking on by default | Set `chat_template_kwargs: {enable_thinking: false}` |
| 4 | Model continued the prompt instead of answering | Used raw `/completion` endpoint | Switched to `/chat/completions` with system/user roles |
| 5 | `IllegalCharacterError` on Excel save (lost a full translation run) | Control chars in extracted PPTX text | Strip `\x00-\x1F` before writing cells |
| 6 | Untranslated ALL-CAPS headers (OBJECTIFS, ACTION, DÉFIS) | SKIP rule treats ≤10-char uppercase as acronyms | Open: fix SKIP logic or catch in review (deferred) |
| 7 | Text inside grouped shapes not extracted | `slide.shapes` does not recurse into GroupShape | Open: add recursive group traversal (next session) |

### Key Takeaways
- A new model architecture needs a matching inference-engine version; downloading the GGUF is not enough.
- Reasoning models ≠ instruct models. Reasoning mode degrades simple tasks like translation; disable it.
- Instruct models must be called via chat templates, not raw completion.
- Inference is **memory-bound**: GPU stays under 50% util because each token requires reading the full model from VRAM. Higher TGP would not help; VRAM capacity/bandwidth is the real constraint. (Training/batched workloads are compute-bound — different story.)
- Persist intermediate artifacts at every stage. One Excel crash already cost a full translation run before checkpointing was added.

### Quality progression
- Deck 1: ~5 clear errors + minor consistency issues over 509 items. Acceptable.
- Deck 2: cleaner; mainly stray foreign-script chars in a name, one untranslated header.
- Deck 3: after adding per-slide context, no spurious mistranslations; remaining gaps were untranslated CAPS headers (SKIP logic) and grouped-shape text (extraction gap), not translation quality.

### Feature: per-slide context (added Day 1)
Each item is now translated with the rest of its slide supplied as reference
context (the item itself excluded, capped at 1500 chars). Goal: consistent
terminology and correct handling of references within a slide. Observed fewer
out-of-context mistranslations on deck 3.

### Hardware Notes
- RTX 5070 Laptop (8GB): Mistral 12B Q4 (~7.3GB) saturates VRAM and throttles throughput.
- Qwen3.5 9B Q4 (~5.68GB) leaves ~2GB headroom; runs cooler (~42°C), full GPU offload at `--gpu-layers 99`.
- CPU usage ~1% at full offload — confirms inference is GPU/memory-bound, not CPU-bound.

### Open Items / Next Session
- [ ] Recursive extraction of grouped shapes (and nested groups) in extractor.py; keep stable IDs for re-insertion
- [ ] Refine SKIP logic so real words in CAPS (OBJECTIFS, ACTION) are translated, true acronyms (NAK, GED) are not
- [ ] Translate decks 3 & 4 to completion; deliver to client
- [ ] Benchmark Mistral for KO→FR direction (per-language-pair model selection)
- [ ] Automatic font resizing for KO text overflow
- [ ] Portfolio prep: .gitignore (exclude input/, 2_translated/, 3_reviewed/, 4_final/, venv/, *.gguf), README (EN/KO/FR), demo GIF
