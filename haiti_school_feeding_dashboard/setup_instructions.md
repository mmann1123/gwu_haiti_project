# Dashboard Setup Instructions

## Prerequisites

Your Google Sheet must have these source sheets:
- `Présence` - Attendance data
- `Taux d'alimentation` - Feeding rate data
- `Info sur les écoles` - School information
- `Comptage Physique` - Physical headcount data
- `Listes prédeterminées` - Dropdown lists

---

## Step 1: Create Hidden Aggregation Sheets

Create these 3 sheets and hide them after setup:

### 1.1 Create "Monthly_Summary" Sheet

This sheet aggregates monthly data by school.

### 1.2 Create "Reason_Counts" Sheet

This sheet counts non-feeding and attendance variation reasons.

### 1.3 Create "School_Current" Sheet

This sheet shows the latest month's data per school.

See `aggregation_sheets.txt` for all formulas.

---

## Step 2: Create Dashboard Sheets

Create 5 new sheets in this order:

1. **Executive Summary** - High-level KPIs
2. **Attendance Analysis** - Attendance trends and reasons
3. **Feeding Rate Analysis** - Feeding days tracking
4. **School Detail** - Individual school drill-down
5. **Supervisor Performance** - Supervisor comparisons

---

## Step 3: Add Named Ranges

Create these named ranges for filter functionality:

| Name | Range | Purpose |
|------|-------|---------|
| `SelectedMonth` | 'Executive Summary'!B2 | Current month filter |
| `SelectedCommune` | 'Attendance Analysis'!B2 | Commune filter |
| `SelectedSchool` | 'School Detail'!B2 | School filter |
| `SelectedSupervisor` | 'Attendance Analysis'!D2 | Supervisor filter |
| `AttendanceTarget` | 'Executive Summary'!E2 | Target attendance rate (e.g., 0.9) |
| `FeedingTarget` | 'Executive Summary'!E3 | Target feeding rate (e.g., 0.9) |

---

## Step 4: Add Data Validation (Dropdowns)

### Month Selector
```
Data validation source: =UNIQUE(TEXT(Présence!A:A,"YYYY-MM"))
```

### Commune Selector
```
Data validation source: ='Info sur les écoles'!C:C
```

### School Selector
```
Data validation source: ='Info sur les écoles'!B:B
```

### Supervisor Selector
```
Data validation source: =UNIQUE('Info sur les écoles'!E:E)
```

---

## Step 5: Add Charts

### Executive Summary Charts

1. **Attendance Gauge Chart**
   - Type: Scorecard or Gauge
   - Data: Current month avg attendance
   - Target: 90%

2. **Feeding Rate Gauge Chart**
   - Type: Scorecard or Gauge
   - Data: Current month avg feeding rate
   - Target: 90%

3. **Regional Comparison Bar Chart**
   - Type: Horizontal bar
   - Data: Avg attendance by commune
   - Sort: Descending

4. **Trend Sparklines**
   - Use SPARKLINE() formula in cells

### Attendance Analysis Charts

1. **Monthly Trend Line Chart**
   - X-axis: Month
   - Y-axis: Attendance rate
   - Multiple series by commune (optional)

2. **Reason Breakdown Pie Chart**
   - Data: Reason counts from Reason_Counts sheet
   - Categories: Cause/effet scolaire, environnemental, etc.

### Feeding Rate Charts

1. **Weekly Feeding Trend**
   - Type: Line chart
   - Data: Weekly feeding rates

2. **Days Fed vs Planned**
   - Type: Stacked bar
   - Series: Actual days, Missing days

3. **Non-Feeding Reasons**
   - Type: Horizontal bar
   - Categories: Communauté, Partenaire, École, Fournisseur, Autre

---

## Step 6: Add Conditional Formatting

### Alert Colors

| Condition | Color | Apply To |
|-----------|-------|----------|
| Attendance < 70% | Red (#EA4335) | Attendance cells |
| Attendance 70-80% | Yellow (#FBBC04) | Attendance cells |
| Attendance > 80% | Green (#34A853) | Attendance cells |
| Feeding rate < 70% | Red (#EA4335) | Feeding rate cells |
| Feeding rate 70-80% | Yellow (#FBBC04) | Feeding rate cells |
| Feeding rate > 80% | Green (#34A853) | Feeding rate cells |

### Change Indicators

For month-over-month changes, use:
- Green up arrow: Positive change
- Red down arrow: Negative change

Formula for arrow:
```
=IF(current-previous>0,"▲ "&TEXT(current-previous,"+0.0%"),"▼ "&TEXT(current-previous,"0.0%"))
```

---

## Step 7: Final Setup

1. **Hide aggregation sheets**
   - Right-click sheet tab > Hide sheet

2. **Protect formulas**
   - Data > Protect sheets and ranges
   - Allow editing only in filter cells

3. **Set default view**
   - View > Freeze > 1 row (for headers)
   - View > Freeze > 1 column (for labels if needed)

4. **Add filter views** (optional)
   - Data > Filter views > Create new filter view
   - Create views for: "My Schools", "Problem Schools", "Monthly Review"
