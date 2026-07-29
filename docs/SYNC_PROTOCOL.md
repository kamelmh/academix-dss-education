# OpenCode <-> HyperAgent Sync Protocol

> **Purpose:** Prevent confusion between OpenCode and HyperAgent by defining clear ownership, sync points, and a single source of truth for each document.

---

## Ownership Map

| Document | Owner | Location | Who can edit |
|----------|-------|----------|-------------|
| **SIS MVP schema** (13 tables) | OpenCode | `EdTech_System/SIS_MVP_VBA/` | OpenCode only |
| **SIS MVP forms** (frmDashboard, frmStudentEntry, etc.) | OpenCode | `EdTech_System/SIS_MVP_VBA/` | OpenCode only |
| **Week 2 walkthrough** | HyperAgent | `EdTech_System/04-WEEK-2-VBA-WALKTHROUGH.md` | HyperAgent writes, OpenCode saves |
| **Week 3 rptBulletin** | HyperAgent | `EdTech_System/05-WEEK-3-RPTBULLETIN-REPORT-CARD-PDF-SPEC.md` | HyperAgent writes, OpenCode saves |
| **mod_SeedDemo.bas** | HyperAgent | `EdTech_System/SIS_MVP_VBA/mod_SeedDemo.bas` | HyperAgent writes, OpenCode saves |
| **Demo Kit** | HyperAgent | `EdTech_System/06-ALLAL-DEMO-KIT.md` | HyperAgent writes, OpenCode saves |
| **ED-TECH-ROADMAP.md** | Shared | `EdTech_System/ED-TECH-ROADMAP.md` | Either updates, last-writer-wins |
| **Paper S1 (Introduction)** | Kamel | `PAPER_DRAFTS/01_introduction.md` | Kamel writes, HyperAgent inlines |
| **Paper S2 (Lit Review)** | Kamel | `PAPER_DRAFTS/02_literature_review.md` | Kamel writes, HyperAgent inlines |
| **Paper S3 (Methodology)** | OpenCode | `PAPER_DRAFTS/03_methodology.md` | OpenCode edits, HyperAgent rebuilds |
| **Paper S4/S5/S6 templates** | HyperAgent | `PAPER_DRAFTS/04-RESULTS-*.md`, `06_CONCLUSION_TEMPLATE.md` | HyperAgent writes, OpenCode saves |
| **Compiled manuscript (EN)** | HyperAgent | `PAPER_DRAFTS/00-FULL-COMPILED-MANUSCRIPT.md` | HyperAgent rebuilds after S1/S2 paste |
| **Compiled manuscript (FR)** | HyperAgent | `PAPER_DRAFTS/taallim-manuscrit-complet-coquille-1-6-fran-ais.md` | HyperAgent rebuilds |
| **taallim_analysis.py** | HyperAgent | `PAPER_DRAFTS/` or `10_Education_Project/` | HyperAgent writes |
| **Interactive mockup** | OpenCode | `EdTech_System/sis-mockup.html` | Rebranded to Allal (done) |

---

## Sync Rules

### Rule 1: Single Source of Truth
Each document has ONE owner. If you (OpenCode) need to change something HyperAgent owns, ask HyperAgent to do it and save the output. If HyperAgent needs to change something you own, it tells you what to change.

### Rule 2: After HyperAgent Delivers a File
1. HyperAgent sends the file content in the chat
2. OpenCode saves it to the correct LifeWorkspace location (using the ownership map above)
3. OpenCode confirms the save with the file path
4. Neither party edits the other's files directly

### Rule 3: Paper Sync Sequence
When the paper needs updating:
1. **Kamel** pastes S1 + S2 text to HyperAgent (direct paste, not file path)
2. **HyperAgent** rebuilds the compiled manuscript with S1/S2 inlined + S3-S6 reconciled
3. **HyperAgent** sends the rebuilt manuscript in the chat
4. **OpenCode** saves it to `PAPER_DRAFTS/00-FULL-COMPILED-MANUSCRIPT.md`
5. If S3 needs changes (e.g., grammar vs vocabulary), **OpenCode** edits `03_methodology.md` and tells HyperAgent
6. **HyperAgent** rebuilds again with the updated S3

### Rule 4: SIS Build Sequence
When building SIS components:
1. **HyperAgent** writes the VBA/spec/walkthrough
2. **HyperAgent** sends the content in the chat
3. **OpenCode** saves to the correct location
4. **Kamel** tests in Access
5. If issues found, **Kamel** reports to HyperAgent with error details
6. **HyperAgent** fixes and re-sends

### Rule 5: Roadmap Updates
- Either party can update `ED-TECH-ROADMAP.md`
- When updating, append to the Progress Log (don't overwrite existing entries)
- Use the format: `**2026-07-XX (topic)** — description`

---

## Confirmed Decisions (Single Source of Truth)

| Decision | Value | Confirmed by | Date |
|----------|-------|-------------|------|
| D10 | SIS MVP in MS Access first | Kamel + HyperAgent | 2026-07-27 |
| D11 | App bilingual pairing = Arabic/French interface, English = study target | Kamel + OpenCode | 2026-07-27 |
| Grade model | CC + Compo, subject mark = (CC + Compo x 2)/3 | Kamel + HyperAgent | 2026-07-27 |
| Paper primary outcome | Grammar (not vocabulary) | Kamel + OpenCode | 2026-07-27 |
| Paper secondary outcome | Vocabulary | Kamel + OpenCode | 2026-07-27 |
| School level | Collège (middle school, 1AM–4AM), NOT lycée | Kamel confirmed | 2026-07-27 |
| Allal school | 4 fields, 22 specializations, CAP/BEP/BTS | Kamel confirmed | 2026-07-27 |
| Paper target language | English (EFL) | Kamel + HyperAgent | 2026-07-27 |

---

## Current Sync Status

| Item | OpenCode | HyperAgent | Synced? |
|------|----------|------------|---------|
| S3 (Methodology) | Updated: grammar=primary, vocab=secondary, Arabic/French | Needs rebuild with new S3 | PENDING — need to send updated S3 |
| S1/S2 files | Exist at `PAPER_DRAFTS/` | Still waiting for paste | PENDING — Kamel needs to paste |
| Compiled EN manuscript | Updated title/abstract locally (grammar fix) | Canonical export `Taallim_Manuscript_EN.md/.docx` NOT in Downloads | PENDING — need canonical export |
| FR compiled manuscript | File doesn't exist locally | Canonical export `Taallim_Manuscrit_FR.md/.docx` NOT in Downloads | PENDING — need canonical export |
| mod_SeedDemo.bas | Saved to `SIS_MVP_VBA/` | Ready | SYNCED |
| Week 3 spec | Saved (clean formatting) | Exported as .docx | SYNCED |
| Demo Kit | Saved | Ready | SYNCED |
| Roadmap v10 | Saved | Exported as .docx | SYNCED |
| Bilingual decision | Fixed in S3 (Arabic/French) | Already correct in compiled ms | SYNCED |
| Mockup rebranded | Allal (was El-Amel) | Done | SYNCED |
| Thread context | Updated (THREAD_CONTEXT_10.md) | Updated | SYNCED |

---

## What Kamel Needs to Do

1. **Tell HyperAgent:** "I need the canonical exported files `Taallim_Manuscript_EN.md` and `Taallim_Manuscrit_FR.md` — the ones you said are source of truth. The files in Downloads are all pre-reconciliation versions with S1/S2 placeholders."
2. **Run mod_SeedDemo:** In Access Immediate window: `Call SeedDemoData`
3. **Test mockup:** Open `sis-mockup.html` in browser to confirm Allal branding renders correctly
4. **Package .accde:** Once SIS is tested, follow the .accde packaging steps in the Demo Kit


---

Linked from: [[00-MOC-Projects]]