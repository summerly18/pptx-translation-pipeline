# Development Log — PPTX Translation Pipeline (FR→KO)

A local LLM pipeline that extracts text from PowerPoint decks, translates FR→KO
with a domain-specific glossary, exports a review spreadsheet, and re-inserts
the approved translation back into the original PPTX.

Stack: Python, python-pptx, openpyxl, llama.cpp (Qwen3.5-9B-Q4_K_M), RTX 5070 Laptop 8GB.

---

## Day 1 — 2026-05-29

### Summary
Built the full pipeline end-to-end and validated it on live client decks
(Archives du Maroc action plans, 4 decks). Restructured into discrete,
resumable stages. Settled model and prompting strategy after several
inference-engine issues.

### Pipeline architecture (final)
Four stages, each writing to its own folder so work survives interruption:

```
input/         원본 PPTX
1_extracted/   추출된 텍스트 (번역 전)   ← 1_extract.py
2_translated/  Qwen 초벌 번역            ← 2_translate.py
3_reviewed/    감수 완료 (Claude+human)
4_final/       번역 삽입된 PPTX          ← reinsert.py
```

### Issues & Resolutions

| # | Issue | Root Cause | Resolution |
|---|-------|-----------|------------|
| 1 | Mistral Nemo 12B: missing translations, English output, prompt leakage | Weak FR→KO coverage; strong only on FR→EN | Switched to Qwen3.5 9B |
| 2 | Qwen3.5 load failure: `missing tensor 'blk.32.ssm_conv1d.weight'` | SSM/hybrid architecture unsupported by installed llama.cpp | Updated llama.cpp to build b9374 |
| 3 | Empty `content`, output only in `reasoning_content` | Qwen3.5 hybrid reasoning model; thinking on by default | Set `chat_template_kwargs: {enable_thinking: false}` |
| 4 | Model continued the prompt instead of answering | Used raw `/completion` endpoint | Switched to `/chat/completions` with system/user roles |
| 5 | `IllegalCharacterError` on Excel save (lost a full translation run) | Control chars in extracted PPTX text | Strip `\x00-\x1F` before writing cells |
| 6 | Untranslated ALL-CAPS headers (OBJECTIFS, ACTION, DÉFIS) | SKIP rule treated ≤10-char uppercase as acronyms | Replaced heuristic with explicit KNOWN_ACRONYMS whitelist (Day 2) |
| 7 | Text inside grouped shapes not extracted | `slide.shapes` does not recurse into GroupShape | Added recursive group traversal (Day 2) |

### Key Takeaways
- A new model architecture needs a matching inference-engine version; downloading the GGUF is not enough.
- Reasoning models ≠ instruct models. Reasoning mode degrades simple tasks like translation; disable it explicitly.
- Instruct models must be called via chat templates, not raw completion endpoints.
- Inference is **memory-bound**: GPU stays under 50% utilization because each token requires reading the full model weights from VRAM. Higher TGP would not help; VRAM capacity/bandwidth is the real bottleneck.
- Persist intermediate artifacts at every stage. One Excel crash already cost a full translation run before checkpointing was added.

### Quality notes (Day 1)
- Deck 1 (509 items): ~5 clear errors + minor consistency issues. Acceptable.
- Deck 2 (151 items): cleaner; mainly stray foreign-script chars in a name, one untranslated header.
- Per-slide context feature added: each item translated with the rest of its slide as reference (self excluded, capped at 1500 chars). Fewer out-of-context mistranslations observed from Deck 3 onward.

### Hardware Notes
- RTX 5070 Laptop (8GB): Mistral 12B Q4 (~7.3GB) saturates VRAM and throttles throughput.
- Qwen3.5 9B Q4 (~5.68GB) leaves ~2GB headroom; runs cooler (~42°C), full GPU offload at `--gpu-layers 99`.
- CPU usage ~1% at full offload — confirms inference is memory-bound, not CPU-bound.

---

## Day 2 — 2026-05-30

### Summary
Fixed grouped-shape extraction, refined SKIP logic, completed translation of
all 4 client decks, ran full code cleanup, wrote README, and published to GitHub.

### Issues & Resolutions

| # | Issue | Root Cause | Resolution |
|---|-------|-----------|------------|
| 1 | Grouped shape text not extracted (decks 3 & 4) | `slide.shapes` flat iteration misses GroupShape children | Recursive traversal; shape_name encodes full path (`Group/Child`) |
| 2 | Nested groups not covered | Single-level fix insufficient for `Group/SubGroup/Child` patterns | Full recursion to arbitrary depth |
| 3 | Table extraction broken after Day 1 refactor | `has_table` check placed inside `has_text_frame` branch | Separated into two independent `if` blocks per shape |
| 4 | ALL-CAPS real words skipped (OBJECTIFS, ACTION, DÉFIS) | Heuristic: uppercase + ≤10 chars → treated as acronym | Replaced with explicit `KNOWN_ACRONYMS` whitelist |
| 5 | Prompt leakage in translation output (1 in ~350 items) | Qwen occasionally outputs meta-commentary before translation | `clean_translation()` post-processor; extracts text after last `번역:` marker |
| 6 | Duplicate Korean in output ("디지털 전환 디지털 전환") | Source deck had Korean annotations embedded in French text | `strip_korean_annotations()` removes Korean runs from mixed FR/KO strings |
| 7 | Email / URL / phone passed to LLM and reformatted | Not in SKIP_PATTERNS | Added regex-based skip for email, URL, phone patterns |
| 8 | `git init` failed — git not installed | Git not on system PATH | Installed Git for Windows; configured user identity |

### Validation results (post-fix)
- Deck 1: `{'textbox': 509}` — confirmed no tables in source file (XML-level check)
- Deck 2: `{'textbox': 63, 'table_cell': 88}` ✅
- Deck 3: `{'textbox': 94, 'table_cell': 20}` ✅ (was 70 items before group fix)
- Deck 4: `{'textbox': 45}` ✅
- Group path matching: 0 mismatches between extractor and reinsert (KHALID deck)
- All `src/*.py` pass `python -m py_compile`

### Workflow division (established Day 2)
- **Claude**: architecture decisions, code review, translation review, study guidance
- **Codex**: implementation of specified tasks
- Rationale: token efficiency; Claude's comparative advantage is judgment and context, not code generation volume.

### Key Takeaways
- GroupShape recursion must be consistent between extractor and reinsert; a naming mismatch fails silently (no error, wrong shapes updated).
- Deck-level context (per-slide) is sufficient for PPTX translation; full-document context would exceed `n_ctx=4096` and is unnecessary given slide independence.
- Validating at XML level (not just python-pptx API) gave definitive proof that Deck 1 has no tables — prevented a false bug chase.

### Publication
- Repository: https://github.com/summerly18/pptx-translation-pipeline
- Version: 1.0.0
- Committed: src/, glossary/, README.md, DEVLOG.md, DECISIONS.md, requirements.txt, .gitignore
- Excluded: input/, 1_extracted/, 2_translated/, 3_reviewed/, 4_final/, venv/, *.gguf

### Open Items / Next
- [ ] Automatic font resizing for KO text overflow in PPTX shapes
- [ ] Benchmark Mistral Nemo for KO→FR direction
- [ ] Batch translation (currently sequential; one API call per item)
- [ ] Next project TBD (RAG pipeline or pipeline extension)
