# Ed-Tech Roadmap — Dual-Track School System + Taallim

*Reference for the unified vision, decisions, and open questions across the Allal Dual-Track School Management System and the Taallim education platform + research paper.*

---

## Executive Summary — One Vision, Two Initiatives

**Mission:** Build *concrete*, AI-powered, bilingual (Arabic / French) education technology for the Algerian market — products that give schools and learners **ownership and lasting value**, not disposable subscription apps.

This roadmap consolidates two threads into a single plan:

| # | Initiative | What it is | Stage (as of Jul 2026) |
|---|---|---|---|
| **1** | **Dual-Track School System** *(working name)* | An AI-powered management platform for Algerian private schools that run **both** academic (elementary → BAC) **and** vocational (التكوين المهني) tracks. Grew out of an "MS Access teaching kit" idea. | MVP feature-complete (Access) |
| **2** | **Taallim** | An education/learning app (flashcards, MCQ, mind-maps) **+** an academic research paper for the *Multilinguales* journal (ASJP). | Components built; paper shell §1–§6 |

**The through-line:** both run on the same engine — AI-generated bilingual educational content — and on the same rare founder profile. Treated together, Taallim can become the *learning-facing module* of the bigger system, and the research paper becomes *sales credibility*.

---

## Progress Log

**2026-07-27 (via OpenCode)** — Parallel build session completed:
- ✅ Literature Review drafted (`02_literature_review.md`, ~1,800 words, 27 references). Paper: §1 ✅ · §2 ✅ · §3–6 pending.
- 📄 New local files: `SCHOOL_SPECIALIZATIONS_TEMPLATE.md`, `00-MASTER-ROADMAP.md`, `MERGED_MASTER.md`.
- 🧭 Decision (D10): SIS MVP built in MS Access first (portable schema for later web migration).
- ⛔ Curriculum Designer blocked until the specializations template is filled.
- 🔧 Workflow: OpenCode = build · HyperAgent = plan/spec · Obsidian *LifeWorkspace* = notes.

**2026-07-27 (later)** — Built an interactive SIS mockup (webpage): 5 clickable screens (Dashboard · Students · Profile · Grade Entry · Bulletin), FR⇄AR toggle with RTL, live recompute, printable Algerian bulletin. Placeholder school "El-Amel" — to be rebranded to Allal.

**2026-07-27 (later still)** — Drafted Methodology §3 (standalone). Quasi-experimental mixed-methods; 120 students (~60/condition); SM-2 flashcards + MCQ + mind-maps; ANCOVA + mixed ANOVA; 14 method refs.

**2026-07-27 (Week 1 build complete — via OpenCode)** — School **confirmed = Allal**. Data locked: **5 academic levels (1AM→BAC)**, **4 vocational fields** (IT, Commerce, Electricity, Mechanics), **22 specializations**, diplomas **CAP/BEP/BTS**. SIS DB built: **13 tables** via `mod_CreateTables.bas` + `mod_Utils.bas` + seed data.
- ✅ Open Question Q1 (specializations) RESOLVED.
- Correction: academic track is middle→secondary (1AM–BAC), not elementary. Schema 11→13 tables (added Field, Diploma, Module, Fee, Config).
- Delivered: Week-2 Forms Build Brief.

**2026-07-27 (paper)** — Created fill-in templates for §4 Results & §5 Discussion (ANCOVA, mixed-ANOVA, correlation/regression, questionnaire + thematic tables; APA stubs; decision rules; effect-size guides).

**2026-07-27 (sync + paper correction)** — Outputs synced to LifeWorkspace. CORRECTION: study **target language = English (EFL)**, not French — updated Methodology §3.

**2026-07-27 (paper §6)** — Drafted §6 Conclusion template. Paper shell complete in structure: §1 ✅ · §2 ✅ · §3 ✅ (draft) · §4/§5/§6 templates.

**2026-07-27 (SIS Week 2)** — 13 tables tested OK. Delivered Week-2 forms walkthrough + paste-ready VBA (frmDashboard, frmStudentEntry, frmClassManager, frmGradeEntry). Flagged: tblGrade CC+Compo vs single Mark?

**2026-07-27 (paper — FR)** — Produced a French translation of §3 Méthodologie for Multilinguales.

**2026-07-27 (paper compile)** — Compiled a single manuscript (§1–§6) inlining §3–§6 with §1/§2 placeholders (Intro/Lit Review never uploaded across attempts). Working title suggested.

**2026-07-27 (paper — analysis script)** — Built & tested `taallim_analysis.py` (pandas + pingouin + statsmodels): ingests the pilot CSV and auto-generates all §4 tables (ANCOVA, mixed ANOVA, correlation/regression, questionnaire) → 9 CSVs + a paste-ready markdown. Ships with a synthetic sample dataset. Verified end-to-end.

**2026-07-27 (paper — FR shell)** — Compiled the French full-manuscript shell (§1–§6). Both EN + FR full shells now exist.

**2026-07-27 (SIS Week 3)** — Delivered the rptBulletin build walkthrough (queries, report layout, TermAverage + ClassRank VBA, frmBulletinPicker with Preview / Export-PDF / Export-whole-class). Demo centrepiece.

**2026-07-27 (paper — title/abstract)** — Drafted 3 title options + a structured abstract skeleton into the English manuscript's Front Matter.

**2026-07-27 (sync+ flag)** — ⚠️ OPEN DECISION: OpenCode's §3 pass changed the *app's* bilingual pairing "Arabic/French" → "Arabic/English" in §3.2 & §3.4. Target language = English (EFL) is settled; the APP interface/gloss pairing needs Kamel's confirmation (Arabic/English vs keep Arabic/French with English only as target). Ripples to abstract, French §3, both manuscripts, and the "bilingual design" subscale. NOT propagated yet.

**2026-07-27 (SIS Week-2 build, interactive)** — Kamel building forms hands-on: frmDashboard = DONE; walked through frmStudentEntry; then → frmClassManager → frmGradeEntry.

**2026-07-27 (SIS Week-3 fix)** — RESOLVED grade-model: tblGrade stores **CC + Compo**. Updated rptBulletin queries to compute subject mark = (CC + Compo×2)/3 (aliased Mark). Report layout + VBA unchanged.

**2026-07-27 (demo prep)** — Built the Allal SIS Demo Kit: pre-demo checklist, ~8-min click-through script, one-page value sheet, likely-Q&A. Flagged: rebrand mockup to Allal.

**2026-07-27 (SIS — MVP feature-complete)** — Walked through frmAttendanceRegister. **MVP feature set complete**: Students · Classes/Enrollment · Grades · Attendance · Bulletin PDF. Next: seed demo data, rehearse the demo, package the .accdb.

---

## Founder Profile & Strategic Edge

Your moat is that you combine four things almost no competitor holds at once:

- **Classroom teacher** — taught at a private school (elementary / middle / secondary).
- **School administrator** — worked *inside* administration, so you know the real pain (enrollment, grades, attendance, fees, Ministry reporting).
- **Certified computing instructor** — 6-month certificate تلقين مبادئ الإعلام الآلي (MS Office), qualifying as *agent administrateur*.
- **Database / VBA builder** — built **Academix DSS** (a VBA decision-support system); you can actually *ship* software.
- **Bilingual (Arabic / French)** — can produce localized content in a market where ed-tech material is overwhelmingly English-only and generic.

**Reference case:** your former school **Allal** — a Ministry-of-Vocational-Training-verified private school that teaches children academically **and** delivers vocational training. It is simultaneously your design partner and your first customer.

---

## Initiative 1 — Dual-Track School Management System

### How the idea evolved (narrow → ambitious)
1. **Start:** *"Is there a market for teaching MS Access — something concrete, not just an app people buy once and forget?"*
2. **Product concept:** an Access Database Starter Kit — ready-to-use `.accdb` databases + tutorials + VBA automation, sold as *owned* assets.
3. **The real vision:** deliver a whole AI-managed system integrating teachers, administration, and students across **both** academic and vocational tracks, with deep curriculum mapping.

### Market findings (2026)
| Signal | Finding |
|---|---|
| Tool relevance | MS Access still in active use in 2026; relational DBs ≈ 57% market share |
| Market size | IT-training market ≈ $91.85B (2025), growing 6.2%/yr |
| Vocational scale (Algeria) | 229 specializations across 20 fields; 291 private vocational schools in Algiers alone |
| Policy tailwind | Gov push for "skills over volume"; 2026 private-school accreditation decree; RNFC-2026 framework |
| Competitor gaps | Udemy = video only · free textbooks = academic · Tes.com = fragmented · McGraw Hill = US-focused · templates = generic. **No one offers a complete, Algerian, bilingual, dual-track system.** |

### Product architecture
**A. School-management core (8 modules):** Students · Enrollment · Grades · Attendance · Fees · Staff · Discipline · Inventory.

**B. Curriculum Designer (the core innovation)** — a 4-level engine:
```
FIELD (20) → SPECIALIZATION (229) → MODULE (2,400+) → COMPETENCY (12,000+)
```
Capabilities: 229 specs pre-loaded · AI generation ("BTS in Cybersecurity, 2 yrs" → modules, competencies, assessment) · RNFC-2026 compliant · academic↔vocational bridges · auto Ministry reports · custom spec builder.

*(For Allal specifically the scope is 4 fields / 22 specializations — the national 20/229 is the full addressable set.)*

### Business model (as scoped so far)
Kit pricing: SIS $50–100 · Inventory Kit $75–150 · School Admin DB $100–200 · CRM $40–80. Master-plan tiers: $35 → $210 (school license). Year-1 target ≈ 240 sales ≈ $15,000. A managed platform likely shifts to per-school licensing (decide — Q8).

---

## Initiative 2 — Taallim Platform & Research Paper

### What's built & tested
`taallim_app.py` (Streamlit, 6 tabs) · `flashcard_system.py` (SM-2) · `mcq_generator.py` · `mindmap_generator.py` · launcher + requirements · `PILOT_EXERCISES.md` (150+) · `REQUIREMENTS_MAPPING.md`.

### The academic paper
- **Target journal:** *Multilinguales* (ASJP).
- **Study:** pilot with 120 students over 6 weeks.
- **Status:** §1 Introduction ✅ · §2 Literature Review ✅ · §3 Methodology ✅ (EFL-aligned draft, EN + FR) · §4/§5/§6 templates ready · analysis script built. Remaining: paste §1/§2 into the compiled manuscript; run the pilot to populate §4–§6.

---

## Shared Foundations & Synergies

| Shared asset | In the School System | In Taallim |
|---|---|---|
| AI content generation | Curriculum / module / lesson generation | Flashcards, MCQs, mind-maps, exercises |
| Bilingual (Ar/Fr) pipeline | Ministry-facing curricula & reports | Learning content & the paper |
| Spaced repetition (SM-2) | Student learning module inside the ERP | Core of the app |
| Pedagogical credibility | Sales proof to schools | The research paper = academic validation |
| Your former school | First ERP customer / design partner | Pilot site for the 120-student study |

> **Strategic play:** make Taallim the learning-facing module of the school system, and use the published paper as third-party validation during school sales.

---

## Decisions Made

| # | Decision | Status |
|---|---|---|
| D1 | Build *concrete*, ownership-based products — not disposable subscription apps | ✅ Firm principle |
| D2 | Target market = Algerian private schools, bilingual Arabic/French | ✅ Firm |
| D3 | Anchor on a dual-track (academic + vocational) system | ✅ Firm |
| D4 | Curriculum Designer is the core differentiator | ✅ Firm |
| D5 | Lead with your former school (Allal) as design partner + first customer | ✅ Firm |
| D6 | Taallim core components built in Python/Streamlit and tested | ✅ Done |
| D7 | Publish a research paper in *Multilinguales* (ASJP) | ✅ Committed |
| D8 | Shift emphasis from "teaching MS Access" → "delivering a managed system" | 🔄 Leaning |
| D9 | Pricing model (one-off kit vs. per-school license) | ❓ Open (Q8) |
| D10 | Build the SIS MVP in MS Access first (portable schema for later web) | ✅ Firm |

---

## Open Questions

**Resolved**
- ✅ Q1 — Allal's specializations: 4 fields, 22 specs, CAP/BEP/BTS.
- ✅ Grade model: tblGrade stores CC + Compo; subject mark = (CC + Compo×2)/3.
- ✅ Paper target language: English (EFL).

**Still open**
- ⚠️ **App bilingual pairing:** Arabic/French (product identity) vs Arabic/English — needs confirmation; ripples through the paper.
- Q3/Q4 — managed AI platform vs owned Access system (a hybrid may reconcile); does it need to *teach* Access at all?
- Q6/Q7 — 2026 accreditation decree requirements; mandatory Ministry reports (format + cadence).
- Q8 — pricing model + Year-1 target under a platform model.
- Q9 — solo-founder sequencing (School System vs Taallim in parallel?).
- Q11 — student-data privacy; where AI calls run (local vs cloud).

---

## Unified Roadmap & Next Steps

**Done:** SIS MVP feature-complete (Students · Classes/Enrollment · Grades · Attendance · Bulletin PDF); paper shell §1–§6 (EN + FR) + analysis script; interactive mockup; demo kit.

**Immediate:** confirm the app's bilingual pairing; rebrand the mockup to Allal; seed realistic demo data; rehearse the demo; package the `.accdb`.

**Phase — Validate:** demo to Allal → secure "we need this" or a change list; run the Taallim pilot (120 students, 6 weeks); populate §4–§6 with real data via the analysis script.

**Phase — Build out & compliance:** remaining modules (Fees, Staff, Discipline, Inventory); RNFC-2026 mapping + Ministry reporting; Curriculum Designer from the 22 specs.

**Phase — Publish & go-to-market:** finish the paper → submit to *Multilinguales*; package system + curriculum; approach Algiers private vocational schools with the published paper as credibility.

---

## Risks & Considerations

| Risk | Note / mitigation |
|---|---|
| **Scope creep** | The vision grew from "Access kit" to "full AI ERP." Ship the SIS wedge first; don't build all 8 modules before validation. |
| **Solo bandwidth** | Two initiatives + a paper is heavy. Sequence: prototype → validate → expand; let the Taallim pilot double as validation. |
| **Regulatory** | The 2026 accreditation decree + Ministry reporting rules may shift requirements — confirm early. |
| **"Concrete vs AI" tension** | Owned/offline (Access) vs. hosted AI platform; a hybrid can satisfy both. |
| **AI reliability** | AI-generated curricula must be checked against RNFC-2026 — human-in-the-loop review. |
| **Single-customer dependency** | Allal is design partner *and* first customer; validate with 1–2 more schools. |
| **Monetization** | $15K Year-1 (kit) is modest; a per-school license may be more sustainable. |


---

Linked from: [[00-MOC-Projects]]