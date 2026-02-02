# Haiti School Feeding Program Dashboard

Complete Google Sheets dashboard implementation for tracking:
- Attendance trends and reasons for absences
- Feeding rates (% of planned days with actual feeding)
- Reasons for non-feeding days

## Files

| File | Description |
|------|-------------|
| `dashboard_formulas_CORRECTED.txt` | **Main file** - All formulas with correct column references |
| `setup_instructions.md` | Step-by-step setup guide |
| `apps_script.gs` | Optional Google Apps Script for automation |
| `aggregation_sheets.txt` | (Legacy) Aggregation sheet formulas |
| `dashboard_formulas.txt` | (Legacy) Original formulas before column mapping |

## Quick Start

### Step 1: Prepare Your Google Sheet
Your data file should have these sheets:
- `Présence` - Monthly attendance data
- `Taux d'alimentation` - Weekly feeding data
- `Info sur les écoles` - School master list
- `Comptage Physique` - Physical headcount data

### Step 2: Create Dashboard Sheets
Create 5 new sheets:
1. **Executive Summary** - High-level KPIs for program managers
2. **Attendance Analysis** - Attendance trends and variation reasons
3. **Feeding Rate Analysis** - Feeding rates and non-feeding reasons
4. **School Detail** - Individual school drill-down
5. **Supervisor Performance** - Supervisor area comparison

### Step 3: Add Formulas
Open `dashboard_formulas_CORRECTED.txt` and copy formulas to each sheet.

### Step 4: Add Charts & Formatting
- Create charts as specified in the formulas file
- Apply conditional formatting (red/yellow/green) for rates

## Column Reference

**Présence Sheet:**
- Col A: Le mois (Month)
- Col B: Nom de l'établissement
- Col E: ID de l'établissement
- Col F: Commune
- Col H: Supervisor
- Col O: Le taux de présence (Attendance rate)
- Col S: La catégorie de variation
- Col T: La raison de la variation

**Taux d'alimentation Sheet:**
- Col A: Semaine commencant
- Col B: Le mois
- Col C: Nom de l'établissement
- Col D: ID de l'établissement
- Col H: Jours prévus
- Col J: Jours alimentés
- Col K: Taux d'alimentation
- Col L, O, R, U, X: Non-feeding categories

## Data Source

Based on: `Summits_MEL_Database_December_2025.xlsx`
