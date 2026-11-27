"""Add the Ascension logo to the title slide (and, optionally, a portrait to the Reflection slide),
then re-validate with IGPPT.

    python add_brand.py <deck.pptx> [--photo portrait.jpg]

Run after:  python <IGPPT>/igppt/scripts/igppt.py build deck.json <deck.pptx>
"""
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
IGPPT = Path(sys.argv[sys.argv.index("--igppt") + 1]) if "--igppt" in sys.argv else HERE.parent / "IGPPT" / "igppt"
sys.path.insert(0, str(IGPPT / "scripts"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402
from ig_kit import DARK, Canvas  # noqa: E402

LOGO = HERE / "assets" / "ascension-logo-reversed.png"   # white wordmark for the black IG title slide


def main():
    deck = Path(sys.argv[1])
    photo = Path(sys.argv[sys.argv.index("--photo") + 1]) if "--photo" in sys.argv else None
    prs = Presentation(str(deck))
    slides = list(prs.slides)

    # Title slide: Ascension logo centred between the subtitle and the date
    w = 4.2
    h = w * 418 / 1829
    Canvas(slides[0], DARK).image(LOGO, (13.333 - w) / 2, 4.65, w=w)

    # Reflection slide (Navy Statement Left): portrait on the right, clear of the quote
    if photo:
        refl = next(s for s in slides if s.slide_layout.name == "Navy Statement Left")
        pic = Canvas(refl, DARK).image(photo, 0, 0, h=7.5)
        pic.left = prs.slide_width - pic.width
        sp = pic._element                      # send behind the text
        sp.getparent().remove(sp)
        refl.shapes._spTree.insert(2, sp)

    prs.save(str(deck))
    words = HERE / "igppt-words.txt"
    cmd = [sys.executable, str(IGPPT / "scripts" / "validate.py"), str(deck), "--spec", str(HERE / "deck.json")]
    if words.exists():
        cmd += ["--words", str(words)]
    sys.exit(subprocess.call(cmd))


if __name__ == "__main__":
    main()
