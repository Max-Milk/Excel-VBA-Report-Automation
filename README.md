# Excel-VBA-Report-Automation
# 📊 Quarterly Report Automation with VBA Form

This project demonstrates how to use **VBA in Excel** with a **UserForm interface** to automate the creation, formatting, and consolidation of quarterly reports into a final yearly summary.

## 📁 Files Included

- `QuarterlyReportForm.xlsm`: Main macro-enabled workbook containing:
  - A custom VBA **UserForm interface**
  - Predefined macros: `AddHeaders`, `FormatData`, `AutoSum`, and `FinalReport`

## 🧰 Features

✅ Create new worksheets  
✅ Navigate to any sheet using a dropdown menu  
✅ Run a fully automated report consolidation process with a single click  
✅ Automatically:
- Add headers to all worksheets
- Format the data with styles and currency formatting
- Calculate totals
- Combine all data into the `"Yearly Report"` sheet

---

## 💻 VBA UserForm Code Explained

The `UserForm` includes a combo box (`cboWhichSheet`) and two buttons (`cmdAddSheet`, `cmdRunReport`):

### `cboWhichSheet_Change`

```vba
Private Sub cboWhichSheet_Change()
    Worksheets(Me.cboWhichSheet.Value).Select
End Sub
````

* Navigates to the selected worksheet from the dropdown list.

### `cmdAddSheet_Click`

```vba
Private Sub cmdAddSheet_Click()
    Worksheets.Add before:=Worksheets(1)
    ActiveSheet.Name = InputBox("Please enter a name for the new worksheet")
End Sub
```

* Adds a new worksheet at the beginning of the workbook and prompts the user to name it.

### `cmdRunReport_Click`

```vba
Private Sub cmdRunReport_Click()
    FinalReport
End Sub
```

* Calls the `FinalReport` procedure to automate the full report workflow.

### `UserForm_Initialize`

```vba
Private Sub UserForm_Initialize()
    Dim i As Integer
    i = 1
    Do While i <= Worksheets.Count
        Me.cboWhichSheet.AddItem Worksheets(i).Name
        i = i + 1
    Loop
End Sub
```

* Populates the dropdown (`cboWhichSheet`) with the names of all worksheets when the form loads.

---

## ⚙️ Final Report Logic (`Module1`)

### `FinalReport`

```vba
Public Sub FinalReport()
    Dim i As Integer
    For i = 1 To Worksheets.Count - 1
        Worksheets(i).Select
        Range("A1").Select
        If ActiveCell.Value <> "" Then
            AutoSum
            AddHeaders
            FormatData

            Range("A2").Select
            Selection.CurrentRegion.Select
            Selection.Copy

            Worksheets("Yearly Report").Select
            Range("A30000").Select
            Selection.End(xlUp).Select
            ActiveCell.Offset(2, 0).Select
            ActiveSheet.Paste
        End If
    Next i
    Columns("C:F").EntireColumn.AutoFit
End Sub
```

* Loops through all worksheets (except "Yearly Report")
* Applies macros (`AutoSum`, `AddHeaders`, `FormatData`)
* Copies data and appends it to `"Yearly Report"`

---

## 🧮 Supporting Macros

### `AutoSum`

```vba
Public Sub AutoSum()
    Dim lastCell As String
    Range("F2").Select
    Selection.End(xlDown).Select
    lastCell = ActiveCell.Address
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "=SUM(F2:" + lastCell + ")"
    ActiveCell.Font.Bold = True
End Sub
```

* Inserts a bolded SUM formula below the last used cell in column **F**.

---

### `AddHeaders`

```vba
Sub AddHeaders()
    Rows("1:1").Insert Shift:=xlDown
    Range("A1").Value = "Division"
    Range("B1").Value = "Category"
    Range("C1").Value = "Jan"
    Range("D1").Value = "Feb"
    Range("E1").Value = "Mar"
    Range("F1").Value = "Total Expense"
End Sub
```

* Adds column headers to row 1 of each worksheet.

---

### `FormatData`

```vba
Sub FormatData()
    Range("A1:F1").Interior.ThemeColor = xlThemeColorAccent1
    Range("A1:F1").Font.Bold = True
    Range("A1:F1").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("C2", Range("C2").End(xlDown).End(xlToRight)).Style = "Currency"
End Sub
```

* Applies background color and bold formatting to headers
* Adds a bottom border
* Formats financial data (Jan–Mar and Total Expense) as currency

---

## ▶️ How to Use the Form

1. Open `QuarterlyReportForm.xlsm`
2. Press `Alt + F11` to view the Visual Basic Editor (optional)
3. Press `Alt + F8` → Select `UserForm1` → Click **Run**
4. Use the form to:

   * Create new sheets
   * Jump between sheets
   * Run the report consolidation

---

## 📧 Contact

* **LinkedIn**: [Max Nguyen Hoang Minh](https://www.linkedin.com/in/max-nguyen-hoang-minh)
* **Email**: [maxnguyenhoangminh@gmail.com](mailto:maxnguyenhoangminh@gmail.com)

```


