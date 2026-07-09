# AGENTS.md

## Cursor Cloud specific instructions

This repository tracks only two things at its root:

- `index.html` — **Product A: "Lipstickk Store System"**, a self-contained, dependency-free client-side web app (Vietnamese UI) for cosmetics pricing/profit calculation. No build step, no backend, no framework. All data persists in the browser via `localStorage` (keys `DB_SI`, `DB_LE`).
- `Archive.zip` — **Product B: "AI-Driven Estimation & Proposal Kit"** (bravesoft), a binary deliverable containing Markdown docs plus two Python generator scripts. It is NOT extracted by default.

### Running Product A (the main web app)

- It is a single static file. Serve it and open in a browser:
  - `python3 -m http.server 8000` (run from the repo root), then open `http://localhost:8000/index.html`.
- Opening `file:///workspace/index.html` directly also works, but serving over HTTP matches how it's used.
- There is nothing to build, lint, or test — no `package.json`, `Makefile`, CI config, or test suite exist. Verification is manual: exercise the 3 tabs (wholesale pricing, retail products, detailed fee analysis) and confirm `localStorage` persistence after reload.
- The product/image thumbnail renders as a broken-image icon when a row is saved without uploading a photo — this is expected app behavior, not a bug.

### Running Product B (the estimation kit inside Archive.zip)

- The `.py` scripts live inside `Archive.zip`; extract it to a scratch dir first (do NOT commit the extracted files — the zip is the source of truth):
  - `unzip -q Archive.zip -d /tmp/estimation-kit`
- Python deps required by the generators: `openpyxl`, `python-pptx`, `Pillow` (installed by the update script). `Pillow` is optional (only used for the white logo on dark PPTX slides; scripts degrade gracefully without it).
- Generate deliverables from the bundled sample input:
  - `python3 /tmp/estimation-kit/05_Tools/generate_estimate_excel.py --input /tmp/estimation-kit/05_Tools/sample_input.json --out /tmp/estimate.xlsx`
  - `python3 /tmp/estimation-kit/05_Tools/generate_proposal_pptx.py --input /tmp/estimation-kit/05_Tools/sample_input.json --out /tmp/proposal.pptx`
- No tests/lint are configured; verification = the scripts produce a valid `.xlsx` / `.pptx`.
