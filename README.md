# pptx-translation-pipeline

**Local LLM pipeline for French → Korean PPTX translation**
로컬 LLM 기반 프랑스어 → 한국어 PPTX 번역 파이프라인

---

## Overview / 개요

A four-stage pipeline that translates French PowerPoint decks into Korean using a locally-hosted LLM, a domain-specific glossary, and a human-in-the-loop review step.

프랑스어 PowerPoint 파일을 로컬 LLM, 도메인 전용 용어집, 인간 감수 단계를 결합하여 한국어로 번역하는 4단계 파이프라인입니다.

Built as a portfolio project demonstrating local LLM pipeline design and domain-specific NLP.


---

## Pipeline Architecture / 파이프라인 구조

```
input/            Original PPTX files / 원본 PPTX
    ↓
1_extract.py  →  1_extracted/    Text extraction (Excel) / 텍스트 추출 (엑셀)
    ↓
2_translate.py →  2_translated/  Qwen first-draft translation / Qwen 초벌 번역
    ↓
  (review)    →  3_reviewed/    Human + Claude review / 인간 + Claude 감수
    ↓
reinsert.py   →  4_final/       Translation inserted into PPTX / 번역 삽입 PPTX
```

Each stage writes to its own folder, so work survives interruption and every artifact is auditable.

각 단계가 별도 폴더에 결과를 저장하므로, 작업이 중단되어도 이어서 진행할 수 있고 모든 산출물을 추적할 수 있습니다.

---

## Key Features / 주요 기능

- **Recursive GroupShape extraction** — Extracts text from nested grouped shapes that `python-pptx` does not traverse by default
- **305-term domain glossary** — Archives & records management terminology (FR↔KO), auto-matched at extraction time
- **Per-slide context** — Each item is translated with the rest of its slide as reference context, improving consistency
- **Resume support** — Translation can be interrupted and resumed; already-translated items are skipped
- **Checkpoint saves** — Progress saved every 10 items to prevent data loss on crash
- **Thinking-mode suppression** — Qwen3.5 is a hybrid reasoning model; `enable_thinking: false` is required to get clean output
- **Post-processing** — Strips prompt leakage, Korean annotations embedded in source text, and skips emails/URLs/phone numbers

---

- **그룹 도형 재귀 추출** — python-pptx 기본 순회로는 접근 불가한 중첩 그룹 도형 내 텍스트 추출
- **305개 도메인 용어집** — 기록관리 분야 전문 용어(FR↔KO) 추출 시 자동 매칭
- **슬라이드 문맥 참조** — 같은 슬라이드의 다른 텍스트를 문맥으로 제공하여 번역 일관성 향상
- **이어하기 지원** — 번역 중단 후 재실행 시 완료된 항목 건너뛰기
- **주기적 중간 저장** — 매 10개마다 저장하여 장애 시 데이터 손실 최소화
- **Thinking 모드 억제** — Qwen3.5는 하이브리드 추론 모델; `enable_thinking: false` 필수
- **후처리** — 프롬프트 누출, 원문 내 한국어 주석, 이메일/URL/전화번호 자동 처리

---

## Tech Stack / 기술 스택

| Component | Choice | Reason |
|-----------|--------|--------|
| LLM inference | llama.cpp b9374 | Local, no API cost, CUDA offload |
| Translation model | Qwen3.5-9B-Q4_K_M | Best FR→KO quality at 8GB VRAM |
| PPTX processing | python-pptx | Read/write with layout preservation |
| Review format | openpyxl (Excel) | Human-readable, easy to edit |
| Hardware | RTX 5070 Laptop 8GB | Full model offload at ~gpu-layers 99 |

### Why Qwen3.5 over Mistral Nemo 12B / Qwen3.5를 선택한 이유

Mistral Nemo 12B was tested first. Issues: missing translations on long sentences, English output instead of Korean, prompt leakage. Root cause: weak FR→KO training coverage.

Qwen3.5 9B resolved all three issues. It also fits in 8GB VRAM with ~2GB headroom (vs Mistral's 7.3GB which saturated VRAM and throttled throughput).

처음에는 Mistral Nemo 12B를 사용했으나, 장문 번역 누락, 영어 출력, 프롬프트 누출 문제가 발생했습니다. 불어→한국어 학습 데이터 부족이 원인이었습니다. Qwen3.5 9B로 교체 후 세 가지 문제 모두 해결됐습니다.

### Why thinking mode must be disabled / Thinking 모드를 꺼야 하는 이유

Qwen3.5 is a hybrid reasoning model. By default, it routes all tasks through its reasoning engine, outputting results in `reasoning_content` while leaving `content` empty. For translation, this wastes tokens on unnecessary chain-of-thought and produces no usable output. Setting `chat_template_kwargs: {enable_thinking: false}` routes output directly to `content`.

Qwen3.5는 하이브리드 추론 모델로, 기본값에서 모든 작업을 추론 엔진으로 처리하여 `reasoning_content`에만 결과를 출력하고 `content`는 비워둡니다. 번역 같은 단순 작업에서는 불필요한 추론 토큰을 낭비합니다. `enable_thinking: false`로 직접 출력 모드를 사용해야 합니다.

---

## Hardware Notes / 하드웨어

- **GPU**: NVIDIA RTX 5070 Laptop (8GB VRAM)
- **`--gpu-layers 99`**: Full model offload to GPU; CPU usage drops to ~1%
- **Inference is memory-bound**: GPU utilization stays under 50% because each token generation requires reading the full model weights from VRAM. Higher TGP would not improve throughput; VRAM bandwidth is the bottleneck.

추론은 memory-bound 작업입니다. 토큰 하나를 생성할 때마다 모델 가중치 전체를 VRAM에서 읽어야 하므로 GPU 연산 유닛이 대기하게 됩니다. TGP를 높여도 속도 향상은 없으며, VRAM 용량과 대역폭이 실질적인 제약입니다.

---

## Installation / 설치

```bash
git clone https://github.com/<your-username>/pptx-translation-pipeline
cd pptx-translation-pipeline
python -m venv venv
# Windows
venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Download Qwen3.5-9B-Q4_K_M.gguf from HuggingFace (bartowski):
```bash
pip install huggingface_hub
huggingface-cli download bartowski/Qwen_Qwen3.5-9B-GGUF \
  --include "Qwen_Qwen3.5-9B-Q4_K_M.gguf" \
  --local-dir /path/to/llama/models
```

---

## Usage / 사용법

### 1. Start the LLM server / LLM 서버 시작
```bash
cd /path/to/llama
./llama-server -m models/Qwen_Qwen3.5-9B-Q4_K_M.gguf --gpu-layers 99 --port 8080
```

### 2. Extract text / 텍스트 추출
```bash
python src/1_extract.py
# Select PPTX → saves to 1_extracted/
```

### 3. Translate / 번역
```bash
python src/2_translate.py
# Select extracted Excel → saves to 2_translated/
# Can be interrupted and resumed
```

### 4. Review / 감수
Open the Excel in `2_translated/`, edit the **최종 번역** column, save to `3_reviewed/`.

`2_translated/`의 엑셀 파일을 열어 **최종 번역** 컬럼을 수정한 후 `3_reviewed/`에 저장합니다.

### 5. Re-insert / 번역 삽입
```bash
python src/reinsert.py
# Select reviewed Excel + original PPTX → saves to 4_final/
```

---

## Project Structure / 프로젝트 구조

```
pptx-translation-pipeline/
├── src/
│   ├── config.py          # Paths, constants, folder setup
│   ├── extractor.py       # PPTX text extraction with group recursion
│   ├── translator.py      # LLM translation via llama.cpp
│   ├── 1_extract.py       # Stage 1 entry point
│   ├── 2_translate.py     # Stage 2 entry point (with resume)
│   └── reinsert.py        # Stage 4 entry point
├── glossary/
│   └── glossary_FR_KO.xlsx  # 305-term FR↔KO glossary
├── input/                 # Source PPTX files (gitignored)
├── 1_extracted/           # Stage 1 output (gitignored)
├── 2_translated/          # Stage 2 output (gitignored)
├── 3_reviewed/            # Stage 3 output (gitignored)
├── 4_final/               # Final translated PPTX (gitignored)
├── requirements.txt
├── DEVLOG.md
└── README.md
```

---

## Limitations & Future Work / 한계 및 향후 계획

- **Image-embedded text**: Text inside images requires OCR (not implemented)
- **Font overflow**: Korean text is longer than French; automatic font resizing not yet implemented
- **KO→FR direction**: Mistral Nemo 12B to be benchmarked for Korean→French translation
- **Batch translation**: Currently sequential (one request per item); batching would improve throughput
- **RAG integration**: Planned as next project — retrieval-augmented generation for domain-specific translation memory

---

- **이미지 내 텍스트**: 이미지 속 텍스트는 OCR 미구현
- **폰트 오버플로우**: 한국어가 불어보다 길어 텍스트박스 넘침 현상 발생 (자동 조정 미구현)
- **한→불 방향**: Mistral Nemo 12B 벤치마킹 예정
- **배치 번역**: 현재 순차 처리; 배치 처리로 속도 개선 가능
- **RAG 통합**: 다음 프로젝트로 계획 중

---

## Development Log / 개발 일지

See [DEVLOG.md](DEVLOG.md) for detailed engineering notes including model selection decisions, inference engine issues, and key learnings.

모델 선택 과정, 추론 엔진 이슈, 핵심 학습 내용 등 상세 엔지니어링 기록은 [DEVLOG.md](DEVLOG.md)를 참조하세요.

---

## Author / 작성자

Contributions, issues, and pull requests are welcome.
GitHub: [summerly18](https://github.com/summerly18)
