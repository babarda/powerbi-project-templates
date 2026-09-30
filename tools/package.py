"""Package each template for download: <slug>-pbip.zip (PBIP project + sample data + how-to),
<slug>-sample_data.xlsx and <slug>-p1.png screenshots, written to dist/ and optionally copied to a site folder.

Run build_pbip.py WITHOUT --local first, so DataFile holds the neutral default path.
Usage: python package.py [site_public_dir]
Author: Badr Abardazzou. License: MIT.
"""
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIST = ROOT / "dist"
SLUGS = {
    "progress-s-curve": "Progress S-Curve",
    "milestone-tracker": "Milestone Tracker",
    "raci-matrix": "RACI Matrix",
    "three-week-lookahead": "3-Week Look-Ahead",
}
HOWTO = """{name}: Power BI template for project controls
Author: Badr Abardazzou. License: MIT. https://babarda.com/tools/powerbi-templates

1. Extract this zip to C:\\PowerBI-Templates\\{slug}\\ (the report looks for
   C:\\PowerBI-Templates\\{slug}\\sample_data.xlsx by default).
2. Replace the sample rows in sample_data.xlsx with yours. Keep sheet names and
   column headers. The README sheet explains every column.
3. Double-click "{name}.pbip". Keep the .pbip, the .Report folder and the
   .SemanticModel folder together.
4. Workbook somewhere else? Home > Transform data > Edit parameters, paste the
   full path of the workbook in DataFile, OK.
5. Home > Refresh.

On first open Power BI may say the model has no data yet: that is normal for a
project file, Refresh loads it.
"""


def main():
    site = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    DIST.mkdir(exist_ok=True)
    for slug, name in SLUGS.items():
        src = ROOT / slug
        zp = DIST / f"{slug}-pbip.zip"
        with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
            z.write(src / f"{name}.pbip", f"{name}.pbip")
            for folder in (f"{name}.Report", f"{name}.SemanticModel"):
                for f in sorted((src / folder).rglob("*")):
                    if f.is_file() and ".pbi" not in f.relative_to(src).parts:
                        z.write(f, f.relative_to(src).as_posix())
            z.write(src / "data" / "sample_data.xlsx", "sample_data.xlsx")
            z.writestr("HOW_TO_USE.txt", HOWTO.format(name=name, slug=slug).replace("\n", "\r\n"))
        shutil.copy(src / "data" / "sample_data.xlsx", DIST / f"{slug}-sample_data.xlsx")
        for png in sorted(src.glob(f"{slug}-p*.png")):
            shutil.copy(png, DIST / png.name)
        print(zp, zp.stat().st_size)
    if site:
        site.mkdir(parents=True, exist_ok=True)
        for f in DIST.iterdir():
            shutil.copy(f, site / f.name)
        print("copied to", site)


if __name__ == "__main__":
    main()
