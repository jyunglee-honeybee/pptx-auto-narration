# 발표자 노트 대본 형식 / Speaker-notes script format

슬라이드 하나의 노트 예시 (One slide's notes):

```
━━ 슬라이드 2 · 클릭 2회 ━━
大家好，今天我们来看一张古代地图。                    ← 슬라이드가 열리면 읽음 / read when the slide opens

▶ 클릭 1   (애니메이션 0.5초)
请看，这就是这张地图的全貌。                          ← 클릭 1 효과와 함께 / with click 1

▶ 클릭 2   (애니메이션 6.0초)
※ 클릭 후 A→B 표시가 자동으로 이어집니다.             ← ※ 줄은 읽지 않음 / not read
【0초】 地图上有两个地名值得注意。
【1.5초·A 표시】 上面的这个，我们叫它A；               ← 클릭 뒤 1.5초 효과에 맞춰 / timed to the effect at +1.5 s
【4.0초·B 표시】 下面的这个，叫B。
```

영어 표기도 됩니다 (English markers work too): `▶ Click 1 (animation 0.5s)`, `【1.5s · A】`.

규칙 / Rules
- `▶ 클릭 N`의 N은 1부터, 그 슬라이드의 클릭 애니메이션 수만큼. / N runs from 1 to the slide's number of click animations.
- 【X초】의 X는 원래 애니메이션에서 그 효과가 클릭 뒤 몇 초에 나오는지(애니메이션 창에서 확인). / X = when that effect starts after the click in the original animation (see the Animation Pane).
- 큐 없는 줄이 큐 문장 뒤에 오면 그 문장에 이어 붙여 읽음. / An un-cued line after a cued one is read as part of it.
- 읽지 않는 줄: `※`로 시작, `━━` 머리줄, `━━ 사용법 ━━` 블록, `(보충` 표시 줄. / Skipped: lines starting with `※`, `━━` headers, a `━━ 사용법 ━━` block, `(보충` marker lines.
