"""Extract TrueType Collection faces so headless renderers can discover them."""

from pathlib import Path

from fontTools.ttLib import TTCollection


fonts_dir = Path(__file__).resolve().parent / "fonts"
source = fonts_dir / "ANGSANA.TTC"
output_dir = fonts_dir / "angsana-extracted"
output_dir.mkdir(exist_ok=True)

for index, font in enumerate(TTCollection(str(source)).fonts):
    font.save(str(output_dir / f"angsana-{index}.ttf"))
