# CAP-LOOP: Closed-Loop Operational Intelligence Cockpit

*CALIBER 2026 — Case 2: Intelligent Manufacturing (PT Chandra Asri Pacific Tbk)*

## Executive summary

CAP-LOOP addresses the **Inaction Gap**: operational alerts and investigations do not consistently become verified, closed corrective actions. In the supplied Case 2 baseline, only **1 of 20 CAPA actions is closed (5% closure; 95% remain open or in progress)**. As of **2 October 2026**, **91 overdue RCAs** carry approximately **US$17.52M** in linked loss. At the same time, the five in-scope assets report availability in the **99.27%–99.82%** range, which can appear healthy while obscuring approximately **US$2.58M in Equipment Performance estimated loss** across those assets.

CAP-LOOP keeps historical backlog separate from forward risk, provides source-linked weekly condition replay, and makes human review and post-repair verification explicit. Incident Database consequences use **Total Loss (actual + potential)** directly; downtime margin is not added a second time.

## Solution pillars

1. **KPI Constitution & Kill-List** — A consistent definition and traceable source lineage for each KPI. Historical overdue-loss backlog is kept separate from replay-based forward risk; untraceable or redundant claims are excluded.
2. **3σ Early Drift Detection** — Robust baselines use the first six eligible normal readings. A signal is flagged at |z| ≥ 3σ for two consecutive weekly readings; the five-asset backtest identifies drift onset in **Weeks 7–8**. This is retrospective evidence, not a guarantee of preventing failure: BL-5702's official alarm occurred in Week 6, one week before the model's drift flag.
3. **Zero-Hardware Motor Energy Load Forecast** — Uses existing hourly `RUN_STATUS=ON` motor-current and plant-rate observations to fit `AMP ~ PLANT_RATE`. It requires no added sensors or hardware. The supplied data has no kWh, voltage, or power-factor basis, so the dashboard does not claim energy savings or calculate an energy cost.
4. **Closed-Loop Action Contracts & Verification** — Human-reviewed, demo-only work-order contracts connect alerts to action ownership. A resolution qualifies after **two consecutive normal post-repair weekly readings with robust z-score ≤ +1.0σ**; negative z-scores are healthy. The supplied CAPA audit has 20 actions: 14 open, 5 in progress, and 1 closed.

## Verified baseline assets

| Equipment tag | Asset | Failure focus | Drift onset | Official alarm | Trip |
|---|---|---|---:|---:|---:|
| PU-2101B | Feed Charge Pump | Mechanical seal leakage | W8 | W15 | W21 |
| KO-3201 | Cracked Gas Compressor | Journal bearing distress / radial vibration | W7 | W10 | W21 |
| PM-4405B | Cooling Water Pump motor | Motor bearing overheating | W7 | W15 | W21 |
| HE-3301 | Feed/Effluent Heat Exchanger | Tube-bundle fouling, duty loss, high differential pressure | W8 | W11 | W21 |
| BL-5702 | Product Blower | Coupling misalignment / high vibration | W7 | W6 | W21 |

Drift is confirmed on the following consecutive weekly reading. The different BL-5702 alarm/drift ordering is intentionally retained rather than presented as a model win.

## Quickstart

Requirements: Python 3.10+ and the packages in `requirements.txt`. The baseline workbooks and RCA presentations must be present under `./data/`.

```bash
python -m pip install -r requirements.txt
python pipeline.py --reveal-all --as-of 2026-10-02
python -m http.server 8000
```

Open <http://localhost:8000/>. The pipeline writes `dashboard_data.json` and `dashboard_data_embedded.js` beside `index.html` using atomic file replacement. `--reveal-all` exports the complete **26-week** condition history for all five assets. The dashboard loads the JSON over HTTP and uses the adjacent embedded snapshot when opened in browser file mode.

To use a different backlog evaluation date, pass `--as-of YYYY-MM-DD`. Optional `--margin-per-h` and `--energy-price-per-kwh` assumptions do not create loss or energy-cost estimates when the source data does not support them.

## GitHub Pages deployment

This is a static dashboard; GitHub Pages does not execute `pipeline.py`. Generate and commit the current JSON and embedded snapshot before publishing.

1. Push the repository's `main` branch to GitHub, including root-level `index.html`, `dashboard_data.json`, and `dashboard_data_embedded.js`.
2. In the repository, open **Settings → Pages**.
3. Under **Build and deployment**, choose **Deploy from a branch**, select `main` and `/(root)`, then save.
4. When the Pages deployment finishes, open the URL shown in Settings → Pages (usually `https://<owner>.github.io/<repository>/`).

The page uses relative references (`./dashboard_data.json` and `./dashboard_data_embedded.js`) and inline SVG illustrations, so it does not depend on a developer workstation path.

## OT cybersecurity boundary

CAP-LOOP is designed for a **Purdue Model Level 3/4 read-only analytics workflow**: source data is batch-exported for analysis and dashboard delivery. The application is not connected to plant control systems and performs **zero write-back to DCS, SAP PM, or CMMS**. Work-order dispatch, confirmations, and sign-off in the web page are local demo interactions only. Engineers remain responsible for validating hypotheses, work scope, and site procedures before action.

## Data scope and interpretation

- Equipment Performance condition histories are weekly and contain 26 readings per asset. Production Data contains separate hourly windows; it is used for the motor-load proxy and is not joined to condition data by week number.
- Consequence is matched to Incident Database **Total Loss (k US$)** (actual plus potential). The forward-risk estimate is a point-in-time replay calculation and must not be added to the historical backlog.
- There are five failure events in this case. Results support operational screening and lead-time analysis, not supervised machine-learning accuracy claims.
- Backlog figures above are explicitly evaluated as of **2026-10-02**. The dashboard's replay cursor defaults to Week 16 while retaining all 26 source weeks.

## Repository contents

- `pipeline.py` — source workbook/RCA processing and atomic JSON export.
- `index.html` — static CAP-LOOP cockpit.
- `dashboard_data.json`, `dashboard_data_embedded.js` — generated dashboard data and file-mode snapshot.
- `data/` — Case 2 baseline source files.
- `backtest_report.csv`, `backtest_report.png` — supporting backtest evidence.
- `requirements.txt`, `.gitignore` — Python dependencies and repository hygiene.

## Initial Git commands

Run these commands from the project root. Replace the example GitHub URL with the repository URL you create:

```bash
git init
git add .
git commit -m "Prepare CAP-LOOP for GitHub Pages"
git branch -M main
git remote add origin https://github.com/<OWNER>/<REPOSITORY>.git
git push -u origin main
```
