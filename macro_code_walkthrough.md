**Walkthrough of the VBA Macro: GenerateMonthlySummaryOrdered**

This document provides a detailed explanation of the `GenerateMonthlySummaryOrdered` macro. The macro is used to process raw financial data from an Excel worksheet ("Sheet1") and generate a cleaned, structured, and formatted report based on specific mapping and exclusion rules. The report is output into another worksheet ("Sheet2").

---

### 1. **Variable and Object Declarations**

```vb
Dim wsSource As Worksheet, wsReport As Worksheet
```

- References to the source and report worksheets.

```vb
Dim lastRow As Long, reportRow As Long
```

- `lastRow`: last row of data in source.
- `reportRow`: tracker for writing to the report sheet.

```vb
Dim dataDict As Object, mapDict As Object
Dim excludeFundDict As Object, excludeAccountDict As Object
Dim fundOrderDict As Object, mapAccountDict As Object
```

- `dataDict`: core data storage for aggregation.
- `mapDict`: maps fund values.
- `excludeFundDict`: funds to ignore.
- `excludeAccountDict`: accounts to ignore.
- `fundOrderDict`: order in which funds should be displayed.
- `mapAccountDict`: maps (fund, parent) to an adjusted parent.

```vb
Dim key As Variant
Dim fund As String, mappedFund As String, desc As String
Dim parent As String, adjParent As String, fiscalYear As String, account As String
Dim periodTotal As Double
Dim monthVal As Variant, monthNum As Integer, monthName As String
```

- Used to capture each relevant data attribute from the source row.

```vb
Dim monthOrder As Variant, monthDict As Object
Dim headers As Collection, allKeys As Collection
```

- `monthOrder`: fixed array of months for ordered display.
- `monthDict`: stores which months are actually present.
- `headers`: column names for the report.
- `allKeys`: stores unique groupings for report rows.

---

### 2. **Sheet Setup**

```vb
Set wsSource = ThisWorkbook.Sheets("Sheet1")
... If wsReport is Nothing ...
```

- Sets up references to "Sheet1" (source) and "Sheet2" (report).
- Clears "Sheet2" if it exists; otherwise, creates it.

---

### 3. **Initialize Data Structures**

```vb
Set dataDict = CreateObject("Scripting.Dictionary")
```

- Initializes all dictionaries used for lookup and grouping.

---

### 4. **Load Reference Data**

#### MappingAccount

```vb
Set wsMapAcc = ThisWorkbook.Sheets("MappingAccount")
```

- Loads mappings from MappingAccount with header.
- Key format: `Fund|Parent`, maps to adjusted parent.

#### MappingFund

```vb
Set wsMap = ThisWorkbook.Sheets("MappingFund")
```

- Maps original fund values to alternate ones.

#### ExcludeFund, ExcludeAccounts

- Funds/accounts listed here are ignored entirely.

#### Order Sheet

```vb
Set wsOrder = ThisWorkbook.Sheets("Order")
```

- Each fund listed is given a display rank (1 = highest priority).
- Others default to a large number (i.e., appear last).

---

### 5. **Process Source Data**

Iterates over rows in Sheet1 to parse and clean values:

```vb
For i = 2 To lastRow
```

#### Fund Mapping

```vb
mappedFund = mapDict(fund) or fund
```

#### Exclusion

```vb
If excludeFundDict.exists(mappedFund) Then GoTo SkipRow
```

#### Adjusted Parent Mapping

```vb
adjParent = Mid(parent, 2) & "00"
If mapAccountDict.exists(fund & "|" & parent) Then
   adjParent = mapAccountDict(...)
```

#### Monthly Aggregation

- Uses a `dataDict(key)` to aggregate monthly values by month.
- Also keeps cumulative `Total`.

#### `key` Construction

```vb
key = mappedFund & "|" & desc & "|" & parent & "|" & fiscalYear
```

- Unique identifier for aggregation row.

---

### 6. **Prepare Column Headers**

```vb
headers.Add "Fund", "Description", ...
```

- Adds fixed columns, then dynamically adds months present in data.

---

### 7. **Sort Keys**

```vb
sortedKeys(i) = allKeys(i)
```

- Converts collection to array.
- Performs bubble sort using `fundOrderDict` rank, then parent value.

---

### 8. **Write Data to Report Sheet**

```vb
wsReport.Cells(reportRow, col).Value = ...
```

- Writes fund, desc, parent, adjusted parent.
- Writes monthly totals and final total.

#### Highlight Negative Totals

```vb
If dict("Total") < 0 Then
  Interior.Color = RGB(255, 199, 206)
```

- Highlights the row in red if the total is negative.

---

### 9. **Format Columns**

```vb
wsReport.Columns(colIndex).NumberFormat = "#,##0.00"
```

- Formats the month and total columns as numbers with commas and 2 decimals.

---

### 10. **Final Steps**

```vb
wsReport.Columns.AutoFit
MsgBox ...
```

- Auto-resizes columns and notifies the user that report is ready.

---

### Summary

The macro:

- Filters and cleans financial source data.
- Maps fund and adjusted parent using external mapping sheets.
- Aggregates by fund/description/parent/FY.
- Displays data in order based on fund list.
- Formats and highlights output for readability.

This document can be used as a technical reference for onboarding new developers or extending the macro’s functionality.

