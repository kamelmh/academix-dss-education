# Academix DSS — Education (Directorate of Education)

Student Information System (SIS) for the Direction de l'Education d'El Bayadh.

## Domain

Education — students, classes, grades, bulletins, report cards.

## Modules

### Shared (7)
`mod_Config`, `mod_Dashboard`, `mod_DemoData`, `mod_StockEngine`, `mod_SupplierRegistry`, `mod_Barcode`, `mod_Reports`

### Education-Specific (3)
| Module | Purpose |
|--------|---------|
| `mod_CreateTables.bas` | Create Access DB tables (13 tables) |
| `mod_SeedDemo.bas` | Seed demo data (students, classes, grades) |
| `mod_Utils.bas` | Utility functions |

## Structure

```
academix-dss-education/
├── modules/           # VBA .bas files
├── docs/              # Build specs, walkthroughs, roadmap
├── mockup/            # SIS HTML mockups
├── CURRICULUM_SEED_DATA.json  # 64 topics across 4 levels
├── ERP_dss_..._education.xlsm # Full DSS workbook
└── README.md
```

## Key Documents

- `docs/SIS_MVP_Build_Spec.md` — Database schema (13 tables)
- `docs/SIS_Week2_VBA_Walkthrough.md` — Forms build guide
- `docs/SIS_Week3_rptBulletin.md` — Report card PDF spec
- `docs/ED-TECH-ROADMAP.md` — 4-phase development plan
- `docs/06-ALLAL-DEMO-KIT.md` — School demo playbook

## Pilot

- **Site:** Allal (private school) + 1 public collège
- **Grades:** 1AM–4AM
- **Students:** ~60 per school

## Dependencies

- Excel 2010+ (VBA host)
- Microsoft Access (for SIS MVP)
- No external libraries

## License

MIT — Mahi Kamel Abdelghani
