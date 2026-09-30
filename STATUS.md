# PowerBI-Project-Templates, status

## Where we are (2026-09-30)
- 4 templates built from source (tools/build_pbip.py, tools/build_data.py): progress-s-curve, milestone-tracker, raci-matrix, three-week-lookahead. PBIR validates with 0 errors.
- Each one opened in Power BI Desktop, refreshed from its data/sample_data.xlsx (TMSL refresh via tools/refresh_model.ps1), screenshotted (tools/render.ps1, tools/crop_shot.py).
- Shipped as PBIP zip (tools/package.py), not .pbit: the Desktop Bridge CLI has no save or export method, and File > Export > Power BI template needs UI control of Desktop.
- Site: barda-ma/src/pages/tools/powerbi-templates.astro plus public/tools/powerbi-templates/ (zips, sample xlsx, PNGs). Not deployed.

## Next action
- Optional .pbit: open each release .pbip in Desktop, File > Export > Power BI template, then set FORMAT = "pbit" in the astro page and copy the .pbit files to public/tools/powerbi-templates/.
- Create the GitHub repo (README, LICENSE, .gitignore are in place).
