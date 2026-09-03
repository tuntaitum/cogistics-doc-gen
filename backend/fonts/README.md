# Fonts

## NotoSansThai-Regular.ttf / NotoSansThai-Bold.ttf

Why: the base-14 PDF fonts (Helvetica etc.) only cover Latin/Western text.
Thai — and any other non-Latin script — renders as solid black boxes with
no error raised, which is easy to miss until you actually look at the
generated PDF. See engine.py's `_ensure_fonts_registered()` for where these
get wired in.

Source: [Noto Sans Thai](https://github.com/google/fonts/tree/main/ofl/notosansthai)
from Google's font repository, licensed under the SIL Open Font License 1.1
(see `OFL.txt` in this folder — required to ship alongside the font).

Google distributes this family as a single **variable** font (one file
covering the whole weight range), not separate static Regular/Bold files.
Reportlab's TTFont embedding works best with static instances, so the two
files here were generated from the variable font using `fonttools`:

```python
from fontTools.ttLib import TTFont as FTFont
from fontTools.varLib.instancer import instantiateVariableFont

for weight, name in [(400, "Regular"), (700, "Bold")]:
    font = FTFont("NotoSansThai-Variable.ttf")
    instantiateVariableFont(font, {"wght": weight, "wdth": 100}, inplace=True)
    font.save(f"NotoSansThai-{name}.ttf")
```

Covers Thai + Latin in one font, so mixed Thai/English text in a single
cell renders correctly without needing per-run font switching. Does **not**
cover CJK (Chinese/Japanese/Korean) — characters outside the font's
coverage are silently omitted by reportlab (not shown as a box), so a cell
with Chinese text would currently lose those characters rather than fail
loudly. Add a CJK font (e.g. Noto Sans SC) the same way if that's ever needed.
