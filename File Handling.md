# 📁 File Handling in VBA – Explanation + Practical Example

## ✅ **What is File Handling in VBA?**

File handling in VBA refers to reading from, writing to, creating, or manipulating files (such as `.txt`, `.csv`, etc.) using VBA code. This is especially useful for automation, logging, importing/exporting data, or interacting with external systems.

## 🔧 Common File Handling Operations in VBA

| Operation         | VBA Method/Function              |
| ----------------- | -------------------------------- |
| Open a file       | `Open` statement                 |
| Read a file       | `Input`, `Line Input`, `Input #` |
| Write to a file   | `Write`, `Print`, `Write #`      |
| Close a file      | `Close #fileNumber`              |
| Check file exists | `Dir` function                   |
| Delete a file     | `Kill` function                  |

## 💡 **Modes for Opening a File in VBA**

```vba
Open filePath For [mode] As #fileNumber
```

| Mode   | Purpose                      |
| ------ | ---------------------------- |
| Input  | Read only                    |
| Output | Write (overwrites if exists) |
| Append | Write (adds to end of file)  |
| Binary | For binary files             |
| Random | For random access files      |

## 📌 Practical Example: Writing and Reading a Log File

### 🔴 Scenario:

You are automating an Excel process and want to:

- **Log some actions or data to a `.txt` file**
- **Read that log later for audit or debugging**

## ✅ Step-by-step VBA Example

```vba
Sub FileHandlingExample()

    Dim filePath As String
    Dim fileNum As Integer
    Dim logText As String
    Dim line As String

    ' Define the log file path
    filePath = ThisWorkbook.Path & "\logfile.txt"

    ' =========================
    ' === Writing to a file ===
    ' =========================
    fileNum = FreeFile ' Get a free file number
    Open filePath For Append As #fileNum
        logText = "Log Entry at " & Now & " - Process started"
        Print #fileNum, logText
    Close #fileNum

    MsgBox "Log written successfully to: " & filePath

    ' =========================
    ' === Reading the file ===
    ' =========================
    fileNum = FreeFile
    Open filePath For Input As #fileNum
        Do Until EOF(fileNum)
            Line Input #fileNum, line
            Debug.Print line  ' Prints to Immediate Window (Ctrl+G to view)
        Loop
    Close #fileNum

End Sub
```

## 🧠 Explanation of Key Parts:

- `FreeFile`: Gets the next available file number.
- `Open ... For Append`: Opens the file to write at the end (does not overwrite).
- `Print #`: Writes a line to the file.
- `Line Input #`: Reads one line at a time.
- `EOF`: End Of File, used to loop through file until the end.
- `Debug.Print`: Useful for reading file output into the Immediate Window.

## 📝 Output:

A file named `logfile.txt` is created in the same folder as the workbook. Example content:

```
Log Entry at 8/5/2025 2:25:53 PM - Process started
```

## ✅ Where this is useful:

- Error or activity logs in macro execution
- Data export/import to/from `.txt` or `.csv`
- Keeping history of changes
- Simple configuration storage

# Practical example

## 📊 Scenario

We have a worksheet named **"SalesData"** with the following **sample data**:

| A (Date)   | B (Product) | C (Region) | D (Quantity) | E (Unit Price) |
| ---------- | ----------- | ---------- | ------------ | -------------- |
| 2025-08-01 | Widget A    | North      | 10           | 25             |
| 2025-08-01 | Widget B    | South      | 5            | 40             |
| 2025-08-02 | Widget A    | East       | 8            | 25             |
| 2025-08-02 | Widget C    | West       | 15           | 60             |
| 2025-08-03 | Widget B    | North      | 12           | 40             |

Each row represents a **sales transaction**.

## 🎯 Goal

We want to generate a **text file summary report** that contains:

- Total sales (Qty × Price) by Product
- Total sales overall

## 📁 Output Example (`SalesSummary.txt`)

```
Sales Summary Report
------------------------

Product-wise Sales:
Widget A: Rs. 450
Widget B: Rs. 680
Widget C: Rs. 900

------------------------
Total Sales: Rs. 2,030
Report generated on: 05-Aug-2025 14:35:21
```

## ✅ VBA Code to Generate This Summary Report

```vba
Sub GenerateSalesSummaryReport()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim product As String
    Dim qty As Double, unitPrice As Double
    Dim totalSale As Double
    Dim salesDict As Object
    Dim filePath As String
    Dim fileNum As Integer
    Dim key As Variant

    Set ws = ThisWorkbook.Sheets("SalesData")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dataRange = ws.Range("A2:E" & lastRow)

    ' Create dictionary to hold product-wise sales
    Set salesDict = CreateObject("Scripting.Dictionary")

    ' Loop through data and calculate sales per product
    For Each row In dataRange.Rows
        product = row.Cells(1, 2).Value
        qty = row.Cells(1, 4).Value
        unitPrice = row.Cells(1, 5).Value
        totalSale = qty * unitPrice

        If salesDict.exists(product) Then
            salesDict(product) = salesDict(product) + totalSale
        Else
            salesDict.Add product, totalSale
        End If
    Next row

    ' Set the output file path (same folder as workbook)
    filePath = ThisWorkbook.Path & "\SalesSummary.txt"
    fileNum = FreeFile

    ' Write to text file
    Open filePath For Output As #fileNum
        Print #fileNum, "Sales Summary Report"
        Print #fileNum, "------------------------"
        Print #fileNum, ""
        Print #fileNum, "Product-wise Sales:"

        totalSale = 0
        For Each key In salesDict.Keys
            Print #fileNum, key & ": Rs. " & Format(salesDict(key), "0.00")
            totalSale = totalSale + salesDict(key)
        Next key

        Print #fileNum, ""
        Print #fileNum, "------------------------"
        Print #fileNum, "Total Sales: Rs. " & Format(totalSale, "0.00")
        Print #fileNum, "Report generated on: " & Format(Now, "dd-mmm-yyyy hh:nn:ss")
    Close #fileNum

    MsgBox "Summary report created at: " & filePath

End Sub
```

## 🧠 Concepts Used:

- **Dictionary**: To store and accumulate sales for each product.
- **File Handling**: `Open ... For Output`, `Print #`, `Close #`
- **Looping through rows** using `.Rows`
- **Formatting output** with `Format()` for currency and datetime

---
