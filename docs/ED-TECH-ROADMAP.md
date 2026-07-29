# Ed-Tech Roadmap — Dual-Track School System + Taallim

## Executive Summary — One Vision, Two Initiatives

**Mission:** Build *concrete*, AI-powered, bilingual (Arabic / French) education technology for the Algerian market — products that give schools and learners **ownership and lasting value**, not disposable subscription apps.

This roadmap consolidates two threads from earlier working sessions into a single plan:

| # | Initiative | What it is | Stage (as of Jul 2026) |
|---|---|---|---|
| **1** | **Dual-Track School System** *(working name)* | An AI-powered management platform for Algerian private schools that run **both** academic (elementary -> BAC) **and** vocational (التكوين المهني) tracks. Grew out of an "MS Access teaching kit" idea. | **MVP feature-complete** (Access) |
| **2** | **Taallim** | An education/learning app (flashcards, MCQ, mind-maps) **+** an academic research paper for the *Multilinguales* journal (ASJP). | Core components built & tested; paper shell complete |

**The through-line:** both run on the same engine — AI-generated bilingual educational content — and on the same rare founder profile (below). Treated together, Taallim can become the *learning-facing module* of the bigger system, and the research paper becomes *sales credibility*.

---

## Progress Log

**2026-07-27 (via OpenCode)** — Parallel build: Literature Review drafted (~1,800 words, 27 refs). Decision D10: SIS MVP built in MS Access first. Curriculum Designer blocked until specializations known. Workflow: OpenCode = build, HyperAgent = plan/spec, Obsidian = notes.

**2026-07-27 (later)** — Built the interactive SIS mockup (5 screens, FR/AR + RTL, printable bulletin). Placeholder school "El-Amel" -> to rebrand to Allal.

**2026-07-27** — Drafted Methodology S3 (standalone); S4/S5 fill-in templates; S6 Conclusion template. Paper shell S1-S6 complete in structure.

**2026-07-27 (Week 1 — OpenCode)** — School confirmed = Allal: 5 levels (1AM->BAC), 4 fields (IT, Commerce, Electricity, Mechanics), 22 specializations, CAP/BEP/BTS. SIS DB = 13 tables via mod_CreateTables.bas + mod_Utils.bas + seed. Q1 (specializations) resolved. Delivered Week-2 Forms Brief.

**2026-07-27 (paper corrections)** — Target language corrected to English (EFL). French S3 translation produced. Compiled EN + FR full manuscripts (with S1/S2 placeholders). Built & tested taallim_analysis.py (auto-generates all S4 tables). Drafted 3 titles + abstract skeleton.

**2026-07-27 (SIS Weeks 2-3)** — Delivered Week-2 forms walkthrough (paste-ready VBA) and Week-3 rptBulletin (queries, report, TermAverage/ClassRank VBA, frmBulletinPicker with PDF export single + whole class). Grade model resolved = CC + Compo -> subject mark (CC + Compo x 2)/3. Walked through all four forms interactively (frmDashboard done). frmAttendanceRegister completed the MVP feature set.

**2026-07-27 (demo prep)** — Built the Allal SIS Demo Kit (checklist, ~8-min click-through script, one-page value sheet, Q&A). Built mod_SeedDemo.bas (3 classes both tracks, ~30 students, CC/Compo grades, attendance).

**2026-07-27 (bilingual decision — RESOLVED)** — App = bilingual Arabic/French interface; English = target/study language. All HyperAgent docs already consistent; only Kamel's local 03_methodology.md needed the S3.2/S3.4 revert (OpenCode).

**2026-07-27 (paper — GROUND TRUTH + reconciliation)** — Actual S1/S2 uploaded. Study outcome = English GRAMMAR (not vocabulary); RQs = AI-integration / grammar-effectiveness / TEACHER perceptions; setting = El Bayadh, 4 secondary schools, ~120 students + 4 teachers; platform multilingual (Ar/Fr/En), offline, curriculum-aligned (64 grammar topics). Rebuilt + re-exported the EN manuscript (inlined S1/S2 + reconciled S3-S6 + merged 27+6 refs). STILL TO SYNC to grammar/teacher: FR manuscript, standalone S3 (EN+FR), S4/S5 templates, analysis-script labels.

**2026-07-27 (housekeeping)** — Exported EN+FR manuscripts, SIS specs, Week-2/Week-3 walkthroughs, roadmap, and analysis script as Markdown + Word. Repaired the literal-\n rendering glitch in this Progress Log and the Week-3 doc.

---

## Founder Profile & Strategic Edge

Your moat is that you combine four things almost no competitor holds at once:

- **Classroom teacher** — taught at a private school (elementary / middle / secondary).
- **School administrator** — worked *inside* administration, so you know the real pain (enrollment, grades, attendance, fees, Ministry reporting).
- **Certified computing instructor** — 6-month certificate تلقين مبادئ الإعلام الآلي (MS Office), qualifying as *agent administrateur*.
- **Database / VBA builder** — built **Academix DSS** (a VBA decision-support system); you can actually *ship* software, not just spec it.
- **Bilingual (Arabic / French)** — can produce localized content in a market where database/ed-tech material is overwhelmingly English-only and generic.

**Reference case:** your former school ("Allal", per prior notes) — a **Ministry-of-Vocational-Training-verified** private school that teaches children academically **and** delivers professional/vocational training. It is simultaneously your design partner and your first customer.

---

## Initiative 1 — Dual-Track School Management System

### How the idea evolved (narrow -> ambitious)
1. **Start:** *"Is there a market for teaching MS Access — something concrete, not just an app people buy once and forget?"*
2. **Product concept:** an **Access Database Starter Kit** — ready-to-use `.accdb` databases + tutorials + VBA automation, sold as *owned* assets.
3. **The real vision (your steer):** stop treating "teaching Access" as the goal and instead **deliver a whole AI-managed system** integrating teachers, administration, and students across **both** academic and vocational tracks, with deep curriculum mapping.

> **Central tension to resolve:** *teaching Access* (a course/kit) vs. *delivering a managed AI platform*. Your latest steer favors the platform — see Decisions (D8) and Open Questions (Q3-Q4).

### Market findings (2026)
| Signal | Finding |
|---|---|
| Tool relevance | MS Access still in active use in 2026; relational DBs = **57%** market share |
| Market size | IT-training market = **$91.85B** (2025), growing **6.2%/yr** |
| Vocational scale (Algeria) | **229** specializations across **20** fields; **291** private vocational schools in Algiers alone |
| Policy tailwind | Gov push for **"skills over volume"**; new **2026 private-school accreditation decree**; **RNFC-2026** framework |
| New 2025-26 specializations | data scientist, 3D graphic design, e-commerce, digital marketing, cybersecurity |
| Competitor gaps | Udemy (57K+ students, 4.8 stars) = video only, free textbooks = academic, no templates, Tes.com = fragmented, McGraw Hill = US-focused, 130+ templates = generic. **No one offers a complete, Algerian, bilingual, dual-track system.** |

### Product architecture
**A. School-management core (8 modules):** Students, Enrollment, Grades, Attendance, Fees, Staff, Discipline, Inventory.

**B. Curriculum Designer (the core innovation)** — a 4-level engine:
```
FIELD (20) -> SPECIALIZATION (229) -> MODULE (2,400+) -> COMPETENCY (12,000+)
```
| Capability | What it does |
|---|---|
| 229 specs pre-loaded | Every Algerian vocational specialization, ready to use |
| AI generation | "BTS in Cybersecurity, 2 yrs" -> modules, competencies, assessment plan; down to weekly lessons |
| RNFC-2026 compliant | New Ministry framework built in |
| Academic <-> vocational bridges | Links e.g. math -> accountancy, physics -> electronics |
| Auto Ministry reports | One click -> compliance document |
| Custom spec builder | Design new specializations for local market demand |

### Business model (as scoped so far)
*Kit/product pricing (original framing):*
| Product | Price |
|---|---|
| Student Information System | $50-100 |
| Inventory Management Kit | $75-150 |
| School Administration DB | $100-200 |
| CRM for Freelancers | $40-80 |

*Master-plan tiers:* 5 tiers, **$35** (basic) -> **$210** (school license). **Year-1 target = 240 sales = $15,000.**

> These numbers reflect the *kit* framing. A managed AI platform likely shifts toward per-school licensing / annual contracts — decide deliberately (Q8).

---

## Initiative 2 — Taallim Platform & Research Paper

### What's built & tested
| File | Purpose | Status |
|---|---|---|
| `taallim_app.py` | Streamlit app, 6 tabs | Ready |
| `Flashcard_System/flashcard_system.py` | SM-2 spaced repetition | Tested |
| `MCQ_Generator/mcq_generator.py` | Multiple-choice generator | Tested |
| `MindMap_Generator/mindmap_generator.py` | Visual concept maps | Tested |
| `run_taallim.bat` | One-click launcher | Ready |
| `requirements_taallim.txt` | Dependencies | Ready |
| `PILOT_EXERCISES.md` | 150+ sample exercises | Ready |
| `REQUIREMENTS_MAPPING.md` | Ministry requirements mapped | Complete |

### The academic paper
- **Target journal:** *Multilinguales* (ASJP — Algerian Scientific Journals Platform).
- **Study outcome:** English GRAMMAR acquisition + teacher perceptions
- **Setting:** El Bayadh, 4 secondary schools, ~120 students + 4 teachers
- **Status:** S1-S6 complete; analysis script built; needs pilot data to populate S4/S5

---

## Shared Foundations & Synergies

| Shared asset | In the School System | In Taallim |
|---|---|---|
| AI content generation | Curriculum / module / lesson generation | Flashcards, MCQs, mind-maps, exercises |
| Bilingual (Ar/Fr) pipeline | Ministry-facing curricula & reports | Learning content & the paper |
| Spaced repetition (SM-2) | Student learning module inside the ERP | Core of the app |
| Pedagogical credibility | Sales proof to schools | The research paper = academic validation |
| Your former school | First ERP customer / design partner | Pilot site for the 120-student study |

> **Strategic play:** make **Taallim the learning-facing module** of the school system, and use the **published paper as third-party validation** during school sales.

---

## Decisions Made

| # | Decision | Status |
|---|---|---|
| D1 | Build *concrete*, ownership-based products — not disposable subscription apps | Firm principle |
| D2 | Target market = Algerian private schools, bilingual Arabic/French | Firm |
| D3 | Anchor on a **dual-track** (academic + vocational) system — a genuine market gap | Firm |
| D4 | Curriculum Designer (229 specs, RNFC-2026, AI-generated) is the core differentiator | Firm |
| D5 | Lead with your former school as design partner + first customer | Firm |
| D6 | Taallim core components built in Python/Streamlit and validated by tests | Done |
| D7 | Publish a research paper in *Multilinguales* (ASJP) for credibility | Committed |
| D8 | Shift emphasis from "teaching MS Access" -> "delivering a managed AI system" | Leaning, not finalized |
| D9 | Pricing model (one-off kit vs. per-school license) | Open (see Q8) |
| D10 | Build the SIS MVP in MS Access first (portable schema for later web) | Firm |
| D11 | App bilingual pairing = Arabic/French interface, English = study target language | Resolved |

---

## Open Questions

**Resolved**
- Q1 — Allal's specializations: 4 fields, 22 specs, CAP/BEP/BTS.
- Grade model: tblGrade stores CC + Compo; subject mark = (CC + Compo x 2)/3.
- Paper target language: English (EFL).
- D11 — App bilingual pairing: Arabic/French interface, English target.

**Still open**
- Q3/Q4 — managed AI platform vs owned Access system (a hybrid may reconcile); does it need to *teach* Access at all?
- Q6/Q7 — 2026 accreditation decree requirements; mandatory Ministry reports (format + cadence).
- Q8 — pricing model + Year-1 target under a platform model.
- Q9 — solo-founder sequencing (School System vs Taallim in parallel?).
- Q11 — student-data privacy; where AI calls run (local vs cloud).

---

## Unified Roadmap & Next Steps

**Done:** SIS MVP feature-complete (Students, Classes/Enrollment, Grades, Attendance, Bulletin PDF); paper shell S1-S6 (EN + FR) + analysis script; interactive mockup; demo kit; mod_SeedDemo.bas; bilingual decision resolved.

**Immediate:** rebrand mockup to Allal; run mod_SeedDemo; rehearse the demo; package the `.accdb`; paste S1/S2 to HyperAgent for full manuscript compilation.

**Phase — Validate:** demo to Allal -> secure "we need this" or a change list; run the Taallim pilot (120 students, 6 weeks); populate S4-S6 with real data via the analysis script.

**Phase — Build out & compliance:** remaining modules (Fees, Staff, Discipline, Inventory); RNFC-2026 mapping + Ministry reporting; Curriculum Designer from the 22 specs.

**Phase — Publish & go-to-market:** finish the paper -> submit to *Multilinguales*; package system + curriculum; approach Algiers private vocational schools with the published paper as credibility.

---

## Risks & Considerations

| Risk | Note / mitigation |
|---|---|
| **Scope creep** | The vision grew from "Access kit" to "full AI ERP." Ship the SIS wedge first; don't build all 8 modules before validation. |
| **Solo bandwidth** | Two initiatives + a paper is heavy for one person. Sequence: prototype SIS -> validate -> expand; let the Taallim pilot double as validation. |
| **Regulatory** | The 2026 accreditation decree + Ministry reporting rules may shift requirements — confirm early (Q6/Q7). |
| **"Concrete vs AI" tension** | Owned/offline (Access) vs. hosted AI platform pull opposite ways; a hybrid (local app + optional AI assist) can satisfy both. |
| **AI reliability** | AI-generated curricula must be checked against RNFC-2026 — keep a human-in-the-loop review step. |
| **Single-customer dependency** | Your former school is design partner *and* first customer; validate with 1-2 more schools before over-fitting. |
| **Monetization** | $15K Year-1 (kit) is modest; a per-school license model may be more sustainable if you choose the platform route. |
