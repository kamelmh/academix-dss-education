# Allal School Profile — CONFIRMED

> **School:** المؤسسة الخاصة علال بدر الدين للتكوين.middleware
> **English:** Private Institution Allal Badr El-Din for Vocational Training
> **Location:** El Bayadh, Algeria
> **Status:** Ministry-of-Vocational-Training-verified
> **Facebook:** facebook.com/IPG.ElBayadh
> **Confirmed:** 2026-07-27

---

## Confirmed Data

### Academic Levels

| Level | French Name | Status |
|-------|-------------|--------|
| 1st Year Middle | 1ère année moyenne (1AM) | ✅ Confirmed |
| 2nd Year Middle | 2ème année moyenne (2AM) | ✅ Confirmed |
| 3rd Year Middle | 3ème année moyenne (3AM) | ✅ Confirmed |
| 4th Year Secondary | 4ème année moyenne (4AM) | ✅ Confirmed |
| Baccalauréat | BAC | ✅ Confirmed |

### Vocational Fields

| Field | French Name | Status |
|-------|-------------|--------|
| IT & Digital | Informatique et Numérique | ✅ Confirmed |
| Commerce | Commerce | ✅ Confirmed |
| Electricity | Électricité | ✅ Confirmed |
| Mechanics | Mécanique | ✅ Confirmed |

### Diploma Levels

| Diploma | French Name | Status |
|---------|-------------|--------|
| CAP | Certificat d'Aptitude Professionnelle | ✅ Confirmed |
| BEP | Brevet d'Études Professionnelles | ✅ Confirmed |
| BTS | Brevet de Technicien Supérieur | ✅ Confirmed |

### Language of Instruction

| Language | Status |
|----------|--------|
| Bilingual (Arabic + French) | ✅ Confirmed |

---

## Curriculum Designer Seed Data

### IT & Digital Specializations

| Code | Name (FR) | Name (AR) | Diploma | Duration |
|------|-----------|-----------|---------|----------|
| IT-01 | Assistant administrateur | مساعد إداري | CAP | 12 months |
| IT-02 | Développeur web | مطور ويب | CAP | 12 months |
| IT-03 | Maintenance informatique | صيانة الحاسوب | CAP | 12 months |
| IT-04 | Intégrateur web | متكامل ويب | BEP | 24 months |
| IT-05 | Administrateur systèmes et réseaux | مسؤول الأنظمة والشبكات | BEP | 24 months |
| IT-06 | Technicien en cybersécurité | تقني في الأمن السيبراني | BTS | 24 months |
| IT-07 | Développeur d'applications | مطور تطبيقات | BTS | 24 months |

### Commerce Specializations

| Code | Name (FR) | Name (AR) | Diploma | Duration |
|------|-----------|-----------|---------|----------|
| COM-01 | Comptable | محاسب | CAP | 12 months |
| COM-02 | Assistant commercial | مساعد تجاري | CAP | 12 months |
| COM-03 | Secrétaire assistant(e) | سكرتير تنفيذي | CAP | 12 months |
| COM-04 | Commerce international | التجارة الدولية | BEP | 24 months |
| COM-05 | Gestion des entreprises | إدارة المؤسسات | BEP | 24 months |
| COM-06 | Marketing et commerce | التسويق والتجارة | BTS | 24 months |

### Electricity Specializations

| Code | Name (FR) | Name (AR) | Diploma | Duration |
|------|-----------|-----------|---------|----------|
| ELEC-01 | Installateur électricien | كهربائي تركيب | CAP | 12 months |
| ELEC-02 | Électricien de maintenance | كهربائي صيانة | CAP | 12 months |
| ELEC-03 | Électrotechnique | الكهرباء الصناعية | BEP | 24 months |
| ELEC-04 | Électronique industrielle | الإلكترونيك الصناعي | BEP | 24 months |
| ELEC-05 | Froid et climatisation | التكييف والتبريد | BTS | 24 months |

### Mechanics Specializations

| Code | Name (FR) | Name (AR) | Diploma | Duration |
|------|-----------|-----------|---------|----------|
| MECH-01 | Mécanique auto | ميكانيكا السيارات | CAP | 12 months |
| MECH-02 | Soudure | اللحام | CAP | 12 months |
| MECH-03 | Mécanique industrielle | الميكانيك الصناعي | BEP | 24 months |
| MECH-04 | Maintenance industrielle | الصيانة الصناعية | BTS | 24 months |

---

## SIS MVP Table Structure

### tblLevel (Academic)
| LevelID | Name_fr | Name_ar | Cycle | OrderIndex |
|---------|---------|---------|-------|------------|
| 1 | 1ère année moyenne | السنة الأولى متوسط | Middle | 1 |
| 2 | 2ème année moyenne | السنة الثانية متوسط | Middle | 2 |
| 3 | 3ème année moyenne | السنة الثالثة متوسط | Middle | 3 |
| 4 | 4ème année moyenne | السنة الرابعة ثانوي | Secondary | 4 |
| 5 | Baccalauréat | البكالوريا | Secondary | 5 |

### tblSpecialization (Vocational)
*(See seed data above — 22 specializations total)*

### tblField (Vocational Fields)
| FieldID | Name_fr | Name_ar |
|---------|---------|---------|
| 1 | Informatique et Numérique | المعلوماتية والرقمية |
| 2 | Commerce | التجارة |
| 3 | Électricité | الكهرباء |
| 4 | Mécanique | الميكانيكا |

---

**Tags:** #allal-school #status/confirmed #curriculum #seed-data #sis #mvp
