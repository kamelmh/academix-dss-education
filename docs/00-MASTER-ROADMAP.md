# Ed-Tech Roadmap — Dual-Track School System + Taallim

> **Source:** HyperAgent (Opus 4.8) — 2026-07-27
> **Status:** Master plan — consolidates both initiatives
> **Local paths:** `10_Education_Project/` + `04_Ideas_&_Projects/MS_Access_School_System/`

---

## Executive Summary — One Vision, Two Initiatives

**Mission:** Build *concrete*, AI-powered, bilingual (Arabic / French) education technology for the Algerian market — products that give schools and learners **ownership and lasting value**, not disposable subscription apps.

| # | Initiative | What it is | Stage (Jul 2026) |
|---|---|---|---|
| **1** | **Dual-Track School System** | AI-powered management platform for Algerian private schools running **both** academic AND vocational tracks | Vision + master plan; prototype not yet built |
| **2** | **Taallim** | Education/learning app (flashcards, MCQ, mind-maps) + academic paper for *Multilinguales* journal (ASJP) | Core components built & tested; paper intro drafted |

**The through-line:** Both run on the same engine — AI-generated bilingual educational content — and on the same rare founder profile. Taallim becomes the *learning-facing module* of the bigger system, and the paper becomes *sales credibility*.

---

## Founder Profile & Strategic Edge

Your moat is that you combine four things almost no competitor holds at once:

- **Classroom teacher** — taught at a private school (elementary / middle / secondary)
- **School administrator** — worked *inside* administration (enrollment, grades, attendance, fees, Ministry reporting)
- **Certified computing instructor** — 6-month certificate تلقين مبادئ الإعلام الآلي (MS Office)
- **Database / VBA builder** — built **Academix DSS**; you can *ship* software, not just spec it
- **Bilingual (Arabic / French)** — localized content in a market where ed-tech material is overwhelmingly English-only

**Reference case:** former school ("Allal") — Ministry-of-Vocational-Training-verified private school teaching children academically AND delivering professional/vocational training. Design partner + first customer.

---

## Initiative 1 — Dual-Track School Management System

### Idea Evolution
1. **Start:** "Is there a market for teaching MS Access — something concrete?"
2. **Product concept:** Access Database Starter Kit — ready-to-use `.accdb` + tutorials + VBA
3. **The real vision:** Deliver a whole AI-managed system integrating teachers, administration, students across both academic and vocational tracks

### Market Findings (2026)

| Signal | Finding |
|---|---|
| Tool relevance | MS Access still active in 2026; relational DBs ~57% market share |
| Market size | IT-training market ~$91.85B (2025), growing 6.2%/yr |
| Vocational scale (Algeria) | **229** specializations across **20** fields; **291** private vocational schools in Algiers alone |
| Policy tailwind | Gov push for "skills over volume"; 2026 accreditation decree; RNFC-2026 |
| New 2025-26 specializations | data scientist, 3D graphic design, e-commerce, digital marketing, cybersecurity |
| Competitor gaps | No one offers a complete, Algerian, bilingual, dual-track system |

### Product Architecture

**A. School-management core (8 modules):**
Students · Enrollment · Grades · Attendance · Fees · Staff · Discipline · Inventory

**B. Curriculum Designer (core innovation):**
```
FIELD (20) → SPECIALIZATION (229) → MODULE (2,400+) → COMPETENCY (12,000+)
```

| Capability | What it does |
|---|---|
| 229 specs pre-loaded | Every Algerian vocational specialization |
| AI generation | "BTS in Cybersecurity, 2 yrs" → modules, competencies, assessment plan |
| RNFC-2026 compliant | New Ministry framework built in |
| Academic ↔ vocational bridges | Links math → accountancy, physics → electronics |
| Auto Ministry reports | One click → compliance document |
| Custom spec builder | Design new specializations for local demand |

### Business Model

| Tier | Price | Includes |
|---|---|---|
| Basic | $35 | Student records only |
| Standard | $75 | + Grades, attendance |
| Professional | $120 | + Fees, staff, reports |
| School | $175 | + All 8 modules |
| Enterprise | $210 | Full system + curriculum designer |

**Year-1 target:** ~240 sales ~ $15,000

---

## Initiative 2 — Taallim Platform & Research Paper

### What's Built

| Component | Status |
|---|---|
| `taallim_app.py` — Streamlit app, 6 tabs | ✅ Ready |
| Flashcard System (SM-2 spaced repetition) | ✅ Tested |
| MCQ Generator | ✅ Tested |
| Mind Map Generator | ✅ Tested |
| `run_taallim.bat` — One-click launcher | ✅ Ready |
| 150+ pilot exercises | ✅ Ready |
| Paper introduction (~950 words) | ✅ Drafted |

### The Academic Paper

- **Target journal:** *Multilinguales* (ASJP)
- **Pilot:** 120 students, 6 weeks
- **Study design:** Quasi-experimental (pre-test/post-test)

---

## Shared Foundations

| Shared asset | School System | Taallim |
|---|---|---|
| AI content generation | Curriculum/module/lesson generation | Flashcards, MCQs, mind-maps |
| Bilingual (Ar/Fr) pipeline | Ministry-facing curricula & reports | Learning content & the paper |
| Spaced repetition (SM-2) | Student learning module | Core of the app |
| Pedagogical credibility | Sales proof to schools | Research paper = validation |
| Former school | First ERP customer | Pilot site for 120-student study |

**Strategic play:** Make Taallim the learning-facing module of the school system. Use the published paper as third-party validation during school sales.

---

## Decisions Made

| # | Decision | Status |
|---|---|---|
| D1 | Build *concrete*, ownership-based products | ✅ Firm |
| D2 | Target market = Algerian private schools, bilingual | ✅ Firm |
| D3 | Anchor on dual-track (academic + vocational) | ✅ Firm |
| D4 | Curriculum Designer (229 specs, RNFC-2026) is core differentiator | ✅ Firm |
| D5 | Former school as design partner + first customer | ✅ Firm |
| D6 | Taallim components built in Python/Streamlit, validated | ✅ Done |
| D7 | Publish paper in *Multilinguales* for credibility | ✅ Committed |
| D8 | Shift from "teaching Access" → "delivering managed AI system" | 🔄 Leaning |
| D9 | Pricing model (one-off vs. per-school license) | ❓ Open |

---

## Open Questions

### Where sessions ended
1. What specializations does former school actually offer?
2. Taallim next step: test app, draft Literature Review, or write Methodology?

### Product & scope
3. Managed AI platform (hosted) vs. desktop/Access-based (offline)?
4. Platform teaches Access or just runs the school?
5. Which module is the MVP wedge? (Likely: Student Information System)

### Market & compliance
6. What does 2026 accreditation decree require?
7. Which Ministry reports are mandatory?
8. Pricing: per-school license vs. per-seat vs. one-off?

### Execution
9. Solo bandwidth — parallel or sequence?
10. Tech stack: Access/VBA or web?
11. Data & privacy: student data handling?

---

## Unified Roadmap

### Immediate (this week)
- [ ] List former school's specializations + academic levels
- [ ] Decide platform vs. owned-system (Q3) and MVP module (Q5)
- [ ] Pick Taallim next step (recommendation: Literature Review)
- [ ] Re-upload MASTER-PLAN.md and CURRICULUM-DESIGNER.md

### Phase 1 (Weeks 1-3) — Prototype
- [ ] Build SIS MVP (Students + Enrollment + Grades + Attendance)
- [ ] Load school's real specializations into Curriculum Designer
- [ ] Produce one auto Ministry report

### Phase 2 (Weeks 3-6) — Validate
- [ ] Demo prototype to former school
- [ ] Run Taallim pilot (120 students, 6 weeks)
- [ ] Draft paper Methodology around pilot

### Phase 3 (Weeks 6-10) — Build out
- [ ] Add remaining modules (Fees, Staff, Discipline, Inventory)
- [ ] Implement RNFC-2026 curriculum mapping
- [ ] Add VBA/AI automation layer
- [ ] Finalize pricing/packaging

### Phase 4 (Weeks 10+) — Publish & go-to-market
- [ ] Collect & analyze pilot data → finish paper → submit
- [ ] Package system + curriculum
- [ ] Approach ~291 Algiers private vocational schools
- [ ] Use published paper as sales credibility

---

## Risks

| Risk | Mitigation |
|---|---|
| Scope creep | Ship SIS wedge first; don't build all 8 before validation |
| Solo bandwidth | Sequence: prototype → validate → expand; pilot doubles as validation |
| Regulatory | 2026 decree + Ministry rules may shift — confirm early |
| Concrete vs. AI tension | Hybrid: local app + optional AI assist |
| AI reliability | Human-in-the-loop review for generated curricula |
| Single-customer dependency | Validate with 1-2 more schools |
| Monetization | $15K Year-1 modest; per-school license may be more sustainable |

---

**Tags:** #hyperagent #roadmap #dual-track #school-system #project/taallim #project/edtech #master-plan
