# pptx-auto-narration

**PowerPoint 발표자 노트의 대본을 AI 음성이 읽고, 클릭 애니메이션도 대본에 맞춰 저절로 넘어가는 발표 파일을 만드는 Claude 스킬**

*A Claude skill that turns a PowerPoint deck with a click-by-click script in its speaker notes into a self-running presentation: Azure neural-voice narration, every click auto-timed to the speech, slides advancing on their own — ready to export as video.*

[한국어](#한국어) · [English](#english)

---

## 한국어

### 무엇을 하나요

- 발표자 노트에 `▶ 클릭 1`, `▶ 클릭 2` … 형식으로 적은 대본을 Microsoft Azure 신경망 음성으로 읽힙니다.
- 원래 클릭으로 넘기던 애니메이션이 음성에 맞춰 자동으로 나타납니다. 한 클릭 안에서 시간차로 나오는 효과도 문장 단위로 맞춥니다.
- 결과물은 슬라이드쇼를 시작하면 끝까지 혼자 진행되는 PPTX입니다. Windows PowerPoint의 '비디오 만들기'로 MP4 영상으로 내보낼 수 있습니다.
- 낭독 전체를 이어 붙인 검토용 MP3와 문장별 시각표도 함께 만들어 줍니다.

### 설치

1. 이 저장소의 **`pptx-auto-narration.zip`** 을 내려받습니다. ZIP은 풀지 마세요. GitHub의 'Code → Download ZIP'으로 받는 저장소 전체 ZIP은 스킬로 올릴 수 없습니다.
2. Claude(claude.ai 또는 데스크톱 앱)에서 **설정 → Capabilities**의 **코드 실행 및 파일 생성**을 켭니다.
3. **Customize → Skills → + → Create skill → Upload a skill**에서 ZIP을 올립니다.

### 준비물

- **Azure Speech 리소스(무료 F0 등급)와 키.** Azure 무료 계정에서 'Speech' 리소스를 만들 때 가격 등급을 **Free F0**로 고릅니다.
- **접속 허용.** Claude 설정 → Capabilities의 허용 도메인에 `<지역>.tts.speech.microsoft.com`을 추가합니다(예: `eastasia.tts.speech.microsoft.com`).
- **Windows PowerPoint.** 재생 확인과 영상 내보내기에 필요합니다. Mac에서 영상으로 내보내면 애니메이션이 빠집니다.

### 사용법

PPT를 첨부하고 Claude에게 이렇게 요청하면 됩니다.

> 이 PPT를 노트 대본을 읽는 자동 발표 파일로 만들어 줘.

Claude가 다음 순서로 진행합니다.
1. 클릭 수와 대본을 점검합니다.
2. Azure 연결을 확인합니다.
3. 목소리 후보 샘플을 만들어 보여 줍니다.
4. 고른 목소리로 전체 음성을 만들고 PPTX를 만듭니다.

노트 형식은 [`references/notes_format.md`](pptx-auto-narration/references/notes_format.md)를 보세요. 노트가 아직 이 형식이 아니면 Claude가 초안을 써서 확인을 받습니다.

### 주의

- Azure 키는 대화에 남습니다. 작업이 끝나면 Azure 포털에서 **Regenerate Key1**로 키를 교체하세요.
- Claude는 소리를 들을 수 없습니다. 발음, 특히 부수 글자와 번체자는 샘플과 검토용 MP3로 직접 확인해야 합니다.
- 무료 등급 한도는 월 50만 자입니다(2026년 11월 기준 Azure 요금표).

---

## English

### What it does

- Reads the script you wrote in the speaker notes (`▶ Click 1`, `▶ Click 2` …) with an Azure neural voice.
- Turns every click animation into a timed effect that fires with the narration. Effects that appear in sequence within one click are aligned sentence by sentence.
- Produces a PPTX that runs start to finish on its own. Export it to MP4 with Windows PowerPoint's *Create a Video*.
- Also writes a review MP3 of the whole narration and a per-sentence timeline.

### Install

1. Download **`pptx-auto-narration.zip`** from this repository. Do not unzip it. GitHub's *Code → Download ZIP* gives the whole repository, which Claude will not accept as a skill.
2. In Claude, turn on **Settings → Capabilities → Code execution and file creation**.
3. Go to **Customize → Skills → + → Create skill → Upload a skill** and upload the ZIP.

### You will need

- **An Azure Speech resource on the Free F0 tier** and its key.
- **Network access.** Add `<region>.tts.speech.microsoft.com` to the allowed domains under Claude's Settings → Capabilities.
- **PowerPoint for Windows** to play the result and export it to video. Exporting video on Mac drops animations.

### Usage

Attach a deck and ask:

> Turn this deck into a self-running presentation that reads the speaker-notes script.

Claude checks the clicks against the script, verifies the Azure connection, offers voice samples, then synthesizes the narration and builds the deck. The notes format is described in [`references/notes_format.md`](pptx-auto-narration/references/notes_format.md); if your notes are not in that format yet, Claude drafts the script for you to approve.

### Notes

- Your Azure key stays in the conversation. Regenerate it in the Azure portal when you are done.
- Claude cannot hear audio, so check the pronunciation yourself with the samples and the review MP3.
- The key technical point: every effect is written as *With Previous* with an absolute delay from the slide start. PowerPoint recomputes *After Previous* chains and ignores their stored delays, so that encoding drifts out of sync.

### Repository layout

```
pptx-auto-narration/          the skill folder (what the ZIP contains)
  SKILL.md                    instructions Claude follows (Korean)
  scripts/narrate_pptx.py     inspect / parse / samples / tts / build
  references/                 notes format and cue-grouping examples
pptx-auto-narration.zip       ready-to-upload skill package
```

The tool can also be run on its own: `python3 pptx-auto-narration/scripts/narrate_pptx.py --help`. It needs Python 3 with lxml, numpy and Pillow, plus ffmpeg.

## License

MIT — see [LICENSE](LICENSE).
