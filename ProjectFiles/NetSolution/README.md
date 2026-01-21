# Alarm Report Generator - Solution Summary

## Overview

This FTOptix NetLogic generates PDF alarm reports from the `EmbeddedDatabase1.AlarmsEventLogger1` table using the **PdfSharp.MigraDoc.Standard** NuGet package.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    alarms_report_manager                        │
├─────────────────────────────────────────────────────────────────┤
│  [ExportMethod] GenerateAlarmsReportWithDates(start, end)       │
│        │                                                        │
│        ▼                                                        │
│  ┌─────────────────┐    ┌───────────────────┐                   │
│  │ FetchAlarmData  │───▶│ EmbeddedDatabase1 │                   │
│  │ (SQL Query)     │◀───│ .AlarmsEventLogger1│                   │
│  └────────┬────────┘    └───────────────────┘                   │
│           │                                                     │
│           ▼                                                     │
│  ┌─────────────────┐    ┌───────────────────┐                   │
│  │ CreateDocument  │───▶│ MigraDoc Document │                   │
│  │ (Header + Body) │    │ (PDF Structure)   │                   │
│  └────────┬────────┘    └───────────────────┘                   │
│           │                                                     │
│           ▼                                                     │
│  ┌─────────────────┐    ┌───────────────────┐                   │
│  │ SaveDocument    │───▶│ ProjectFiles/     │                   │
│  │ (PDF Render)    │    │ Reports/*.pdf     │                   │
│  └─────────────────┘    └───────────────────┘                   │
└─────────────────────────────────────────────────────────────────┘
```

---

## Key Components

### 1. NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| PdfSharp.MigraDoc.Standard | 1.51.15 | PDF document generation |
| System.Drawing.Common | 10.0.2 | Required dependency for PDFsharp |

### 2. Type Aliases (Namespace Conflict Resolution)

FTOptix and MigraDoc share common type names. Aliases resolve conflicts:

```csharp
using MigraDocColor = MigraDoc.DocumentObjectModel.Color;
using MigraDocTable = MigraDoc.DocumentObjectModel.Tables.Table;
using MigraDocRow = MigraDoc.DocumentObjectModel.Tables.Row;
using MigraDocCell = MigraDoc.DocumentObjectModel.Tables.Cell;
using MigraDocColumn = MigraDoc.DocumentObjectModel.Tables.Column;
using MigraDocVerticalAlignment = MigraDoc.DocumentObjectModel.Tables.VerticalAlignment;
```

### 3. Runtime Initialization

Required in `Start()` for .NET Core/.NET 5+ compatibility:

```csharp
AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);
Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
```

---

## Data Flow

### Query Construction

The database uses ISO 8601 date format with **T separator**:
```
2026-01-21T14:32:10.1762615
```

Query format:
```sql
SELECT * FROM AlarmsEventLogger1 
WHERE LocalTime >= "2026-01-01T00:00:00" AND LocalTime <= "2026-01-31T23:59:59"
ORDER BY LocalTime DESC
```

### Locale-Aware Message Columns

The alarm table stores localized messages in columns like:
- `Message_it-IT` (Italian)
- `Message_en-US` (English)

The solution reads `Session.ActualLocaleId` to select the correct column:

```csharp
string currentLocale = Session.ActualLocaleId;  // e.g., "it-IT"
string localizedMessageColumn = $"Message_{currentLocale}";  // "Message_it-IT"
```

If the localized column doesn't exist, it falls back to the generic `Message` column.

---

## PDF Report Structure

```
┌────────────────────────────────────────────────────────────┐
│ [Customer Logo]      ALARM REPORT           [Last Logo]   │  ◄── Header Row
├────────────────────────────────────────────────────────────┤
│ Machine Name:      │ My Machine Name                      │  ◄── Info Table
│ Filter Start Date: │ 2026-01-01 00:00:00                  │
│ Filter End Date:   │ 2026-01-31 23:59:59                  │
├────────────────────────────────────────────────────────────┤
│                    Alarm History                           │  ◄── Section Title
├────────────────────────────────────────────────────────────┤
│ Activation Date/Time │ Alarm Text │ Associated Device      │  ◄── Table Header
├──────────────────────┼────────────┼────────────────────────┤
│ 2026-01-21 14:32:10  │ High Temp  │ Sensor_01              │  ◄── Data Rows
│ 2026-01-21 14:30:05  │ Low Press  │ Valve_02               │      (alternating
│ ...                  │ ...        │ ...                    │       colors)
├────────────────────────────────────────────────────────────┤
│ Total records: 25                                          │
├────────────────────────────────────────────────────────────┤
│ Generated on 2026-01-22 10:15:00 - Page 1 of 2            │  ◄── Footer
└────────────────────────────────────────────────────────────┘
```

---

## Exported Methods

### GenerateAlarmsReportWithDates(DateTime startDate, DateTime endDate)

Main method to generate the PDF report.

**Usage from FTOptix UI:**
```
Call: alarms_report_manager.GenerateAlarmsReportWithDates
Parameters: 
  - startDate: DateTimePicker1.Value
  - endDate: DateTimePicker2.Value
```

### SetCustomerLogo(string logoPath)
Sets the customer logo (left side of header).

### SetLastLogo(string logoPath)
Sets the Last logo (right side of header).

### SetMachineName(string machineName)
Sets the machine name displayed in the report.

---

## Configuration

### Customizable Colors

Modify these fields in the class to change the report appearance:

```csharp
// Header colors
private readonly MigraDocColor HeaderBackgroundColor = new MigraDocColor(0, 51, 102);
private readonly MigraDocColor HeaderTextColor = MigraDocColors.White;

// Table header colors
private readonly MigraDocColor TableHeaderBackgroundColor = new MigraDocColor(0, 102, 153);
private readonly MigraDocColor TableHeaderTextColor = MigraDocColors.White;

// Alternating row colors
private readonly MigraDocColor TableRowEvenColor = new MigraDocColor(240, 248, 255);
private readonly MigraDocColor TableRowOddColor = MigraDocColors.White;

// Border color
private readonly MigraDocColor TableBorderColor = new MigraDocColor(180, 180, 180);
```

### NetLogic Variables

Optional variables that can be added to the NetLogic:

| Variable | Type | Default | Description |
|----------|------|---------|-------------|
| MachineName | String | "Machine Name" | Displayed in report header |
| CustomerLogoPath | String | "%PROJECTDIR%/imgs/logo1.svg" | Path to customer logo |
| LastLogoPath | String | "%PROJECTDIR%/imgs/logo2.svg" | Path to Last logo |
| OutputPath | String | (auto-generated) | Custom PDF output path |

---

## Output

PDFs are saved to: `ProjectFiles/Reports/`

Filename format:
```
AlarmReport_{StartDate}_{EndDate}_{Timestamp}.pdf
```

Example:
```
AlarmReport_20260101_20260131_20260122101500.pdf
```

---

## Troubleshooting

| Issue | Cause | Solution |
|-------|-------|----------|
| No alarms returned | Date format mismatch | Use ISO 8601 format with T separator |
| Encoding 1252 error | Missing encoding provider | Call `Encoding.RegisterProvider()` in Start() |
| System.Drawing error | .NET Core compatibility | Call `AppContext.SetSwitch()` in Start() |
| Color/Table conflicts | Namespace collision | Use type aliases (MigraDocColor, etc.) |
| Wrong language in text | Locale column not found | Check column names match `Message_{locale}` pattern |
