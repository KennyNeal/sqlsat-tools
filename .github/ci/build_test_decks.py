"""CI smoke test for the slide-deck builders.

Builds both decks (generate_slide_template.py, generate_raffle_deck.py) from
synthetic manifests with placeholder images, then asserts the parts of the
output packages that the builders construct by hand: the .potx content-type
flip, and the raffle deck's sections, custom shows, and loop flag.
"""
import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

from PIL import Image

REPO = Path(__file__).resolve().parents[2]
SCRIPTS = REPO / "scripts"


def make_images(work):
    specs = [
        ("logo.png", (400, 300)),
        ("s1.png", (300, 200)),
        ("s2.png", (200, 300)),
        ("s3.png", (250, 250)),
        ("qr.png", (200, 200)),
        ("brug.png", (300, 90)),
    ]
    for name, size in specs:
        Image.new("RGB", size, (1, 49, 105)).save(work / name)


def run_builder(script, manifest_path):
    subprocess.run(
        [sys.executable, str(SCRIPTS / script), str(manifest_path)],
        check=True,
        cwd=SCRIPTS,
    )


def main():
    work = Path(tempfile.mkdtemp(prefix="sqlsat-ci-"))
    make_images(work)

    groups = [
        {"tier": "global", "title": "Global Sponsor",
         "sponsors": [{"name": "Globex", "logoPath": str(work / "s1.png")}]},
        {"tier": "gold", "title": "Gold Sponsors",
         "sponsors": [{"name": "Initech", "logoPath": str(work / "s2.png")},
                      {"name": "Umbrella", "logoPath": str(work / "s3.png")}]},
    ]
    common = {
        "eventName": "Day of Data CI",
        "footerText": "#CITag",
        "primaryColor": "013169",
        "secondaryColor": "F7C15D",
        "logoPath": str(work / "logo.png"),
    }

    slide_manifest = dict(
        common,
        brugLogoPath=str(work / "brug.png"),
        sponsors=[s for g in groups for s in g["sponsors"]],
        outputPath=str(work / "SlideTemplate.potx"),
    )
    raffle_manifest = dict(
        common,
        qrPath=str(work / "qr.png"),
        evalUrl="https://example.sessionize.com/",
        loopAdvanceSeconds=8,
        maxPerGridSlide=8,
        individualTiers=["global"],
        gridGroups=[["gold"]],
        heroTiers=["global", "gold"],
        excludeSponsors=["Umbrella"],
        groups=groups,
        outputPath=str(work / "RaffleDeck.pptx"),
    )

    slide_json = work / "slide_manifest.json"
    raffle_json = work / "raffle_manifest.json"
    slide_json.write_text(json.dumps(slide_manifest))
    raffle_json.write_text(json.dumps(raffle_manifest))

    run_builder("generate_slide_template.py", slide_json)
    run_builder("generate_raffle_deck.py", raffle_json)

    with zipfile.ZipFile(work / "SlideTemplate.potx") as z:
        ct = z.read("[Content_Types].xml")
        assert b"template.main+xml" in ct, ".potx content-type flip missing"

    with zipfile.ZipFile(work / "RaffleDeck.pptx") as z:
        pres = z.read("ppt/presentation.xml")
        assert b"custShowLst" in pres, "custom show list missing"
        assert b'name="Recognition"' in pres, "Recognition show/section missing"
        assert b'name="Raffle"' in pres, "Raffle show/section missing"
        assert b"sectionLst" in pres, "PowerPoint sections missing"
        props = z.read("ppt/presProps.xml")
        assert b'loop="1"' in props, "loop-until-Esc flag missing"

    print("Both decks built and verified.")


if __name__ == "__main__":
    main()
