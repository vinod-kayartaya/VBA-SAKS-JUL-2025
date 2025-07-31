# VBA 101 Complete Self-Learning Tutorial

_Foundations + Intermediate Level_

## 📊 Sample Dataset: Fashion Sales Data

Throughout this tutorial, we'll work with a fashion retail dataset containing the following columns:

- **Column A**: Order_ID (e.g., ORD001, ORD002)
- **Column B**: Customer_Name (e.g., Sarah Johnson, Mike Chen)
- **Column C**: Product_Category (e.g., Dresses, Shoes, Accessories)
- **Column D**: Product_Name (e.g., Summer Dress, Running Shoes)
- **Column E**: Size (e.g., S, M, L, XL, 8, 9, 10)
- **Column F**: Color (e.g., Red, Blue, Black)
- **Column G**: Quantity (e.g., 1, 2, 3)
- **Column H**: Unit_Price (e.g., 45.99, 89.50)
- **Column I**: Total_Amount (e.g., 45.99, 179.00)
- **Column J**: Order_Date (e.g., 2024-01-15, 2024-02-20)
- **Column K**: Sales_Rep (e.g., Alice Smith, Bob Wilson)

---

# 🔹 Session 1: VBA Introduction, Macros, Variables & Operators

## 1.1 Introduction to VBA and Developer Tools

### What is VBA?

Visual Basic for Applications (VBA) is a programming language built into Microsoft Office applications. It allows you to automate repetitive tasks, create custom functions, and build interactive applications within Excel.

### Enabling Developer Tools

1. Go to **File** → **Options** → **Customize Ribbon**
2. Check the **Developer** checkbox on the right panel
3. Click **OK**

The Developer tab will now appear in your ribbon with options like:

- **Visual Basic**: Opens the VBA editor
- **Macros**: View and run existing macros
- **Record Macro**: Start recording your actions

### The Visual Basic Editor (VBE)

Access the VBE by pressing `Alt + F11` or clicking **Developer** → **Visual Basic**.

**Key Components:**

- **Project Explorer**: Shows all open workbooks and their VBA components
- **Properties Window**: Displays properties of selected objects
- **Code Window**: Where you write your VBA code
- **Immediate Window**: For testing code snippets (View → Immediate Window)

## 1.2 Recording & Editing Macros

### Recording Your First Macro

Let's record a macro that formats our fashion sales data headers:

1. Click **Developer** → **Record Macro**
2. Name: `FormatHeaders`
3. Description: `Formats the header row of fashion sales data`
4. Click **OK**
5. Perform these actions:
   - Select row 1 (headers)
   - Make text bold (`Ctrl + B`)
   - Apply blue background color
   - Center align the text
6. Click **Developer** → **Stop Recording**

### Viewing the Recorded Code

Press `Alt + F11` to open VBE and look at the generated code:

```vb
Sub FormatHeaders()
'
' FormatHeaders Macro
' Formats the header row of fashion sales data
'
    Rows("1:1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Color = RGB(173, 216, 230)
        .Pattern = xlSolid
    End With
    With Selection
        .HorizontalAlignment = xlCenter
    End With
End Sub
```

### Understanding the Code Structure

- `Sub` and `End Sub`: Define the beginning and end of a procedure
- Comments: Lines starting with `'` are comments (ignored during execution)
- `Selection`: Refers to currently selected cells
- `With...End With`: Groups multiple operations on the same object

## 1.3 Variables and Data Types

### What are Variables?

Variables are containers that store data values. They make your code more flexible and readable.

### Declaring Variables

Use the `Dim` statement to declare variables:

```vb
Dim variableName As DataType
```

### Common Data Types

| Data Type | Description                       | Example        |
| --------- | --------------------------------- | -------------- |
| `String`  | Text data                         | "Summer Dress" |
| `Integer` | Whole numbers (-32,768 to 32,767) | 5              |
| `Long`    | Large whole numbers               | 1000000        |
| `Double`  | Decimal numbers                   | 45.99          |
| `Boolean` | True/False values                 | True           |
| `Date`    | Date and time values              | #1/15/2024#    |
| `Variant` | Any data type (default)           | Any value      |

### Variable Examples with Fashion Data

```vb
Sub VariableExamples()
    ' Declare variables for fashion sales data
    Dim customerName As String
    Dim productCategory As String
    Dim quantity As Integer
    Dim unitPrice As Double
    Dim orderDate As Date
    Dim isOnSale As Boolean

    ' Assign values
    customerName = "Sarah Johnson"
    productCategory = "Dresses"
    quantity = 2
    unitPrice = 45.99
    orderDate = #1/15/2024#
    isOnSale = True

    ' Display in Immediate Window (Ctrl + G to view)
    Debug.Print "Customer: " & customerName
    Debug.Print "Category: " & productCategory
    Debug.Print "Quantity: " & quantity
    Debug.Print "Unit Price: $" & unitPrice
    Debug.Print "Order Date: " & orderDate
    Debug.Print "On Sale: " & isOnSale
End Sub
```

### Variable Naming Rules

- Must start with a letter
- Cannot contain spaces (use underscores: `unit_price`)
- Cannot be VBA keywords (`Sub`, `End`, `If`, etc.)
- Case-insensitive but use consistent capitalization

## 1.4 Arithmetic & Logical Operators

### Arithmetic Operators

| Operator | Description         | Example        |
| -------- | ------------------- | -------------- |
| `+`      | Addition            | `5 + 3 = 8`    |
| `-`      | Subtraction         | `10 - 4 = 6`   |
| `*`      | Multiplication      | `6 * 2 = 12`   |
| `/`      | Division            | `15 / 3 = 5`   |
| `\`      | Integer Division    | `15 \ 4 = 3`   |
| `Mod`    | Modulus (remainder) | `15 Mod 4 = 3` |
| `^`      | Exponentiation      | `2 ^ 3 = 8`    |

### Logical Operators

| Operator | Description                 | Example                      |
| -------- | --------------------------- | ---------------------------- |
| `AND`    | Both conditions true        | `(5 > 3) AND (2 < 4)` = True |
| `OR`     | At least one condition true | `(5 < 3) OR (2 < 4)` = True  |
| `NOT`    | Opposite of condition       | `NOT (5 > 3)` = False        |

### Comparison Operators

| Operator | Description           | Example                  |
| -------- | --------------------- | ------------------------ |
| `=`      | Equal to              | `"Red" = "Red"` = True   |
| `<>`     | Not equal to          | `"Red" <> "Blue"` = True |
| `>`      | Greater than          | `10 > 5` = True          |
| `<`      | Less than             | `3 < 7` = True           |
| `>=`     | Greater than or equal | `5 >= 5` = True          |
| `<=`     | Less than or equal    | `4 <= 6` = True          |

### Practical Example: Calculate Fashion Sales

```vb
Sub CalculateSales()
    Dim quantity As Integer
    Dim unitPrice As Double
    Dim discount As Double
    Dim totalAmount As Double
    Dim finalAmount As Double
    Dim taxRate As Double

    ' Sample fashion sale data
    quantity = 3
    unitPrice = 29.99
    discount = 0.1  ' 10% discount
    taxRate = 0.08  ' 8% tax

    ' Calculate total before discount
    totalAmount = quantity * unitPrice

    ' Apply discount
    finalAmount = totalAmount - (totalAmount * discount)

    ' Add tax
    finalAmount = finalAmount + (finalAmount * taxRate)

    ' Display results
    Debug.Print "Quantity: " & quantity
    Debug.Print "Unit Price: $" & unitPrice
    Debug.Print "Subtotal: $" & totalAmount
    Debug.Print "After 10% discount: $" & (totalAmount - (totalAmount * discount))
    Debug.Print "Final Amount (with tax): $" & Round(finalAmount, 2)

    ' Check if bulk order (quantity > 2)
    If quantity > 2 Then
        Debug.Print "Bulk order - eligible for additional discount!"
    End If
End Sub
```

### String Concatenation

Use the `&` operator to join strings:

```vb
Sub StringExample()
    Dim firstName As String
    Dim lastName As String
    Dim fullName As String

    firstName = "Sarah"
    lastName = "Johnson"
    fullName = firstName & " " & lastName

    Debug.Print "Customer: " & fullName
End Sub
```

---

# 🔹 Session 2: Control Flow & User Interaction

## 2.1 Conditional Statements

### If...Then...Else Statement

**Syntax:**

```vb
If condition Then
    ' Code to execute if condition is True
ElseIf another_condition Then
    ' Code to execute if another_condition is True
Else
    ' Code to execute if all conditions are False
End If
```

### Fashion Sales Example: Size Category

```vb
Sub CategorizeSizes()
    Dim size As String
    Dim sizeCategory As String

    ' Get size from cell B2
    size = Range("E2").Value

    ' Categorize the size
    If size = "XS" Or size = "S" Then
        sizeCategory = "Small"
    ElseIf size = "M" Or size = "L" Then
        sizeCategory = "Medium"
    ElseIf size = "XL" Or size = "XXL" Then
        sizeCategory = "Large"
    Else
        sizeCategory = "Unknown"
    End If

    ' Display result
    MsgBox "Size " & size & " is categorized as: " & sizeCategory
End Sub
```

### Select Case Statement

More efficient for multiple conditions:

**Syntax:**

```vb
Select Case expression
    Case value1
        ' Code for value1
    Case value2, value3
        ' Code for value2 or value3
    Case Is > value4
        ' Code for values greater than value4
    Case Else
        ' Default code
End Select
```

### Fashion Example: Product Category Discount

```vb
Sub ApplyDiscountByCategory()
    Dim category As String
    Dim discountRate As Double
    Dim unitPrice As Double
    Dim finalPrice As Double

    ' Get data from specific cells
    category = Range("C2").Value  ' Product Category
    unitPrice = Range("H2").Value  ' Unit Price

    ' Apply discount based on category
    Select Case category
        Case "Dresses", "Tops"
            discountRate = 0.15  ' 15% discount
        Case "Shoes"
            discountRate = 0.1   ' 10% discount
        Case "Accessories"
            discountRate = 0.05  ' 5% discount
        Case Else
            discountRate = 0     ' No discount
    End Select

    ' Calculate final price
    finalPrice = unitPrice * (1 - discountRate)

    ' Update the cell with new price
    Range("H2").Value = finalPrice

    ' Show message
    MsgBox category & " discount applied: " & discountRate * 100 & "%" & vbCrLf & _
           "New price: $" & Round(finalPrice, 2)
End Sub
```

## 2.2 Looping Constructs

### For...Next Loop

Used when you know how many times to repeat:

**Syntax:**

```vb
For counter = start_value To end_value Step increment
    ' Code to repeat
Next counter
```

### Example: Process 10 Fashion Orders

```vb
Sub ProcessTenOrders()
    Dim i As Integer
    Dim orderID As String
    Dim customerName As String

    ' Process orders in rows 2 to 11
    For i = 2 To 11
        orderID = Range("A" & i).Value
        customerName = Range("B" & i).Value

        ' Add processing timestamp in column L
        Range("L" & i).Value = "Processed: " & Now()

        Debug.Print "Processed Order " & orderID & " for " & customerName
    Next i

    MsgBox "10 orders processed successfully!"
End Sub
```

### For Each Loop

Used to iterate through collections:

```vb
Sub HighlightSaleItems()
    Dim cell As Range
    Dim productRange As Range

    ' Define the range of products (assuming data in rows 2-100)
    Set productRange = Range("D2:D100")

    ' Loop through each cell in the range
    For Each cell In productRange
        ' Highlight cells containing "Sale"
        If InStr(cell.Value, "Sale") > 0 Then
            cell.Interior.Color = RGB(255, 255, 0)  ' Yellow background
            cell.Font.Bold = True
        End If
    Next cell

    MsgBox "Sale items highlighted!"
End Sub
```

### Do While Loop

Continues while condition is true:

**Syntax:**

```vb
Do While condition
    ' Code to repeat
Loop
```

### Example: Find Last Row with Data

```vb
Sub FindLastOrderRow()
    Dim rowNum As Integer

    rowNum = 2  ' Start from row 2 (after headers)

    ' Continue while there's data in column A (Order_ID)
    Do While Range("A" & rowNum).Value <> ""
        Debug.Print "Row " & rowNum & ": " & Range("A" & rowNum).Value
        rowNum = rowNum + 1
    Loop

    MsgBox "Last row with data: " & (rowNum - 1)
End Sub
```

### Do Until Loop

Continues until condition becomes true:

```vb
Sub ProcessUntilTarget()
    Dim rowNum As Integer
    Dim totalSales As Double
    Dim targetAmount As Double

    rowNum = 2
    totalSales = 0
    targetAmount = 1000  ' Target sales amount

    ' Process orders until we reach target sales
    Do Until totalSales >= targetAmount
        totalSales = totalSales + Range("I" & rowNum).Value
        Debug.Print "Row " & rowNum & " - Running Total: $" & totalSales
        rowNum = rowNum + 1

        ' Safety check to avoid infinite loop
        If rowNum > 1000 Then Exit Do
    Loop

    MsgBox "Target reached! Total sales: $" & totalSales
End Sub
```

## 2.3 Message Boxes and Input Boxes

### Message Boxes

Display information to users:

**Syntax:**

```vb
MsgBox prompt, [buttons], [title], [helpfile], [context]
```

### Message Box Examples

```vb
Sub MessageBoxExamples()
    Dim customerName As String
    Dim orderTotal As Double
    Dim response As Integer

    customerName = "Sarah Johnson"
    orderTotal = 156.99

    ' Simple message
    MsgBox "Order processed successfully!"

    ' Message with title
    MsgBox "Welcome to Fashion Store!", , "System Message"

    ' Information message
    MsgBox "Customer: " & customerName & vbCrLf & _
           "Order Total: $" & orderTotal, vbInformation, "Order Summary"

    ' Yes/No question
    response = MsgBox("Apply 10% discount for bulk order?", vbYesNo + vbQuestion, "Discount Option")

    If response = vbYes Then
        MsgBox "Discount applied!"
    Else
        MsgBox "No discount applied."
    End If
End Sub
```

### Message Box Button Types

| Constant        | Description                 |
| --------------- | --------------------------- |
| `vbOKOnly`      | OK button only (default)    |
| `vbOKCancel`    | OK and Cancel buttons       |
| `vbYesNo`       | Yes and No buttons          |
| `vbYesNoCancel` | Yes, No, and Cancel buttons |
| `vbRetryCancel` | Retry and Cancel buttons    |

### Message Box Icon Types

| Constant        | Description          |
| --------------- | -------------------- |
| `vbCritical`    | Critical (X) icon    |
| `vbQuestion`    | Question (?) icon    |
| `vbExclamation` | Warning (!) icon     |
| `vbInformation` | Information (i) icon |

### Input Boxes

Get input from users:

**Syntax:**

```vb
InputBox(prompt, [title], [default], [xpos], [ypos], [helpfile], [context])
```

### Input Box Examples

```vb
Sub InputBoxExamples()
    Dim customerName As String
    Dim discountPercent As Double
    Dim orderDate As String

    ' Get customer name
    customerName = InputBox("Enter customer name:", "Customer Information", "Enter name here")

    If customerName <> "" Then
        ' Get discount percentage
        discountPercent = InputBox("Enter discount percentage (0-50):", "Discount", "0")

        ' Validate discount
        If IsNumeric(discountPercent) And discountPercent >= 0 And discountPercent <= 50 Then
            ' Get order date
            orderDate = InputBox("Enter order date (MM/DD/YYYY):", "Order Date", Format(Date, "mm/dd/yyyy"))

            ' Display summary
            MsgBox "Order Details:" & vbCrLf & _
                   "Customer: " & customerName & vbCrLf & _
                   "Discount: " & discountPercent & "%" & vbCrLf & _
                   "Date: " & orderDate, vbInformation, "Order Summary"
        Else
            MsgBox "Invalid discount percentage!", vbExclamation
        End If
    Else
        MsgBox "Operation cancelled.", vbInformation
    End If
End Sub
```

### Interactive Fashion Sales Entry

```vb
Sub InteractiveSalesEntry()
    Dim lastRow As Integer
    Dim customerName As String
    Dim productName As String
    Dim quantity As Integer
    Dim unitPrice As Double
    Dim totalAmount As Double
    Dim continueEntry As Integer

    Do
        ' Get input from user
        customerName = InputBox("Enter customer name:", "New Sale Entry")
        If customerName = "" Then Exit Do  ' User clicked Cancel

        productName = InputBox("Enter product name:", "New Sale Entry")
        If productName = "" Then Exit Do

        quantity = InputBox("Enter quantity:", "New Sale Entry", "1")
        If Not IsNumeric(quantity) Then
            MsgBox "Invalid quantity!", vbExclamation
            Exit Do
        End If

        unitPrice = InputBox("Enter unit price:", "New Sale Entry", "0.00")
        If Not IsNumeric(unitPrice) Then
            MsgBox "Invalid price!", vbExclamation
            Exit Do
        End If

        ' Calculate total
        totalAmount = quantity * unitPrice

        ' Find next empty row
        lastRow = Range("A" & Rows.Count).End(xlUp).Row + 1

        ' Add data to worksheet
        Range("A" & lastRow).Value = "ORD" & Format(lastRow - 1, "000")  ' Order ID
        Range("B" & lastRow).Value = customerName
        Range("D" & lastRow).Value = productName
        Range("G" & lastRow).Value = quantity
        Range("H" & lastRow).Value = unitPrice
        Range("I" & lastRow).Value = totalAmount
        Range("J" & lastRow).Value = Date  ' Order date

        ' Ask if user wants to continue
        continueEntry = MsgBox("Sale entry added successfully!" & vbCrLf & _
                              "Add another sale?", vbYesNo + vbQuestion, "Continue?")

    Loop While continueEntry = vbYes

    MsgBox "Sales entry completed!", vbInformation
End Sub
```

---

# 🔹 Session 3: Working with Excel Objects, Ranges & Data

## 3.1 Excel Object Model

### Understanding the Hierarchy

Excel follows an object hierarchy:

- **Application** (Excel itself)
  - **Workbooks** (Collection of all open workbooks)
    - **Workbook** (Individual Excel file)
      - **Worksheets** (Collection of all sheets)
        - **Worksheet** (Individual sheet)
          - **Range** (Cells or cell ranges)
            - **Cell** (Individual cell)

### Application Object

Represents Excel itself:

```vb
Sub ApplicationExamples()
    ' Display Excel version
    MsgBox "Excel Version: " & Application.Version

    ' Turn off screen updating for faster execution
    Application.ScreenUpdating = False

    ' Your code here...

    ' Turn screen updating back on
    Application.ScreenUpdating = True

    ' Turn off alerts temporarily
    Application.DisplayAlerts = False
    ' Delete something or save file
    Application.DisplayAlerts = True
End Sub
```

## 3.2 Working with Workbooks

### Opening and Creating Workbooks

```vb
Sub WorkbookOperations()
    Dim wb As Workbook
    Dim filePath As String

    ' Create new workbook
    Set wb = Workbooks.Add
    wb.SaveAs "C:\Fashion_Sales_Report.xlsx"

    ' Open existing workbook
    filePath = "C:\Fashion_Data.xlsx"
    Set wb = Workbooks.Open(filePath)

    ' Reference current workbook
    Set wb = ThisWorkbook

    ' Reference active workbook
    Set wb = ActiveWorkbook

    ' Close workbook
    wb.Close SaveChanges:=True
End Sub
```

### Working with Multiple Workbooks

```vb
Sub CopyDataBetweenWorkbooks()
    Dim sourceWB As Workbook
    Dim targetWB As Workbook

    ' Create target workbook for reports
    Set targetWB = Workbooks.Add

    ' Open source workbook with fashion data
    Set sourceWB = Workbooks.Open("C:\Fashion_Sales_Data.xlsx")

    ' Copy headers from source to target
    sourceWB.Sheets("Sales").Range("A1:K1").Copy
    targetWB.Sheets("Sheet1").Range("A1").PasteSpecial xlPasteAll

    ' Copy first 10 records
    sourceWB.Sheets("Sales").Range("A2:K11").Copy
    targetWB.Sheets("Sheet1").Range("A2").PasteSpecial xlPasteAll

    ' Clear clipboard
    Application.CutCopyMode = False

    ' Save target workbook
    targetWB.SaveAs "C:\Fashion_Top10_Report.xlsx"

    MsgBox "Data copied successfully!"
End Sub
```

## 3.3 Working with Worksheets

### Accessing Worksheets

```vb
Sub WorksheetOperations()
    Dim ws As Worksheet

    ' Reference by index (first sheet)
    Set ws = Worksheets(1)

    ' Reference by name
    Set ws = Worksheets("Sales_Data")

    ' Reference active sheet
    Set ws = ActiveSheet

    ' Add new worksheet
    Set ws = Worksheets.Add
    ws.Name = "Monthly_Report"

    ' Move worksheet to end
    ws.Move After:=Worksheets(Worksheets.Count)
End Sub
```

### Worksheet Management Example

```vb
Sub ManageFashionSheets()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Integer

    ' Array of sheet names for fashion business
    sheetNames = Array("Sales_Data", "Inventory", "Customers", "Reports")

    ' Create sheets if they don't exist
    For i = 0 To UBound(sheetNames)
        ' Check if sheet exists
        Dim sheetExists As Boolean
        sheetExists = False

        For Each ws In Worksheets
            If ws.Name = sheetNames(i) Then
                sheetExists = True
                Exit For
            End If
        Next ws

        ' Create sheet if it doesn't exist
        If Not sheetExists Then
            Set ws = Worksheets.Add
            ws.Name = sheetNames(i)

            ' Add headers based on sheet type
            Select Case sheetNames(i)
                Case "Sales_Data"
                    ws.Range("A1:K1").Value = Array("Order_ID", "Customer_Name", "Product_Category", _
                                                   "Product_Name", "Size", "Color", "Quantity", _
                                                   "Unit_Price", "Total_Amount", "Order_Date", "Sales_Rep")
                Case "Inventory"
                    ws.Range("A1:F1").Value = Array("Product_ID", "Product_Name", "Category", _
                                                   "Size", "Color", "Stock_Quantity")
                Case "Customers"
                    ws.Range("A1:E1").Value = Array("Customer_ID", "Customer_Name", "Email", _
                                                   "Phone", "Total_Orders")
            End Select

            ' Format headers
            ws.Range("1:1").Font.Bold = True
            ws.Range("1:1").Interior.Color = RGB(173, 216, 230)
        End If
    Next i

    MsgBox "Fashion business worksheets created/verified!"
End Sub
```

## 3.4 Working with Ranges and Cells

### Range References

```vb
Sub RangeReferences()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Single cell
    ws.Range("A1").Value = "Order_ID"

    ' Multiple cells (same value)
    ws.Range("A1:C1").Value = "Header"

    ' Different syntax for same range
    ws.Cells(1, 1).Value = "Order_ID"  ' Row 1, Column 1 (A1)
    ws.Cells(1, 2).Value = "Customer"  ' Row 1, Column 2 (B1)

    ' Range using Cells
    ws.Range(ws.Cells(1, 1), ws.Cells(1, 11)).Font.Bold = True

    ' Named ranges
    ws.Range("A1:K1").Name = "HeaderRow"
    Range("HeaderRow").Interior.Color = RGB(200, 200, 200)
End Sub
```

### Dynamic Range Selection

```vb
Sub WorkWithDynamicRanges()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim dataRange As Range

    Set ws = ActiveSheet

    ' Find last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Find last column with data in row 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' Create dynamic range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Work with the range
    MsgBox "Data range: " & dataRange.Address & vbCrLf & _
           "Rows: " & lastRow & vbCrLf & _
           "Columns: " & lastCol

    ' Apply formatting to entire data range
    dataRange.Borders.LineStyle = xlContinuous
End Sub
```

## 3.5 Cell and Range Manipulations

### Reading and Writing Data

```vb
Sub ReadWriteFashionData()
    Dim ws As Worksheet
    Dim i As Long
    Dim orderID As String
    Dim customerName As String
    Dim totalAmount As Double
    Dim grandTotal As Double

    Set ws = Worksheets("Sales_Data")
    grandTotal = 0

    ' Read data from worksheet
    For i = 2 To 11  ' Assuming data in rows 2-11
        orderID = ws.Cells(i, 1).Value      ' Column A
        customerName = ws.Cells(i, 2).Value  ' Column B
        totalAmount = ws.Cells(i, 9).Value   ' Column I

        ' Add to grand total
        grandTotal = grandTotal + totalAmount

        ' Write status to column L
        ws.Cells(i, 12).Value = "Processed"

        Debug.Print "Order: " & orderID & " | Customer: " & customerName & " | Amount: $" & totalAmount
    Next i

    ' Write grand total to summary cell
    ws.Cells(1, 13).Value = "Grand Total"
    ws.Cells(2, 13).Value = grandTotal
    ws.Cells(2, 13).NumberFormat = "$#,##0.00"

    MsgBox "Grand Total: $" & Format(grandTotal, "#,##0.00")
End Sub
```

### Formatting Ranges

```vb
Sub FormatFashionReport()
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim dataRange As Range
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Format headers
    Set headerRange = ws.Range("A1:K1")
    With headerRange
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)  ' White text
        .Interior.Color = RGB(0, 102, 204)  ' Blue background
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With

    ' Format data area
    Set dataRange = ws.Range("A2:K" & lastRow)
    With dataRange
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Format currency columns (Unit_Price and Total_Amount)
    ws.Range("H2:I" & lastRow).NumberFormat = "$#,##0.00"

    ' Format date column
    ws.Range("J2:J" & lastRow).NumberFormat = "mm/dd/yyyy"

    ' Auto-fit columns
    ws.Columns("A:K").AutoFit

    ' Add alternating row colors
    For i = 2 To lastRow
        If i Mod 2 = 0 Then  ' Even rows
            ws.Range("A" & i & ":K" & i).Interior.Color = RGB(242, 242, 242)
        End If
    Next i

    MsgBox "Fashion report formatted successfully!"
End Sub
```

## 3.6 Looping Through Data

### Loop Through Rows

```vb
Sub AnalyzeSalesByCategory()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim category As String
    Dim totalAmount As Double

    ' Counters for each category
    Dim dressCount As Long, dressTotal As Double
    Dim shoeCount As Long, shoeTotal As Double
    Dim accessoryCount As Long, accessoryTotal As Double

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Initialize counters
    dressCount = 0: dressTotal = 0
    shoeCount = 0: shoeTotal = 0
    accessoryCount = 0: accessoryTotal = 0

    ' Loop through each row of data
    For i = 2 To lastRow
        category = ws.Cells(i, 3).Value     ' Product_Category (Column C)
        totalAmount = ws.Cells(i, 9).Value  ' Total_Amount (Column I)

        ' Categorize and count
        Select Case UCase(category)
            Case "DRESSES", "TOPS", "BOTTOMS"
                dressCount = dressCount + 1
                dressTotal = dressTotal + totalAmount
            Case "SHOES", "BOOTS", "SNEAKERS"
                shoeCount = shoeCount + 1
                shoeTotal = shoeTotal + totalAmount
            Case "ACCESSORIES", "BAGS", "JEWELRY"
                accessoryCount = accessoryCount + 1
                accessoryTotal = accessoryTotal + totalAmount
        End Select
    Next i

    ' Display results
    Dim report As String
    report = "FASHION SALES ANALYSIS" & vbCrLf & vbCrLf
    report = report & "Clothing: " & dressCount & " items, $" & Format(dressTotal, "#,##0.00") & vbCrLf
    report = report & "Shoes: " & shoeCount & " items, $" & Format(shoeTotal, "#,##0.00") & vbCrLf
    report = report & "Accessories: " & accessoryCount & " items, $" & Format(accessoryTotal, "#,##0.00") & vbCrLf
    report = report & vbCrLf & "Total: " & (dressCount + shoeCount + accessoryCount) & " items, $" & _
             Format(dressTotal + shoeTotal + accessoryTotal, "#,##0.00")

    MsgBox report, vbInformation, "Sales Analysis"
End Sub
```

### Loop Through Columns

```vb
Sub ValidateDataColumns()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim i As Long, j As Long
    Dim cellValue As Variant
    Dim errorCount As Long
    Dim errorList As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = 11  ' Columns A to K
    errorCount = 0

    ' Loop through each row
    For i = 2 To lastRow
        ' Loop through each column
        For j = 1 To lastCol
            cellValue = ws.Cells(i, j).Value

            ' Check for empty required fields
            If j <= 4 And cellValue = "" Then  ' Order_ID, Customer_Name, Category, Product_Name required
                errorCount = errorCount + 1
                errorList = errorList & "Row " & i & ", Column " & Split("A,B,C,D,E,F,G,H,I,J,K", ",")(j - 1) & ": Empty required field" & vbCrLf
            End If

            ' Validate specific columns
            Select Case j
                Case 7  ' Quantity (Column G)
                    If Not IsNumeric(cellValue) Or cellValue <= 0 Then
                        errorCount = errorCount + 1
                        errorList = errorList & "Row " & i & ", Column G: Invalid quantity" & vbCrLf
                    End If
                Case 8, 9  ' Unit_Price (H) and Total_Amount (I)
                    If Not IsNumeric(cellValue) Or cellValue < 0 Then
                        errorCount = errorCount + 1
                        errorList = errorList & "Row " & i & ", Column " & IIf(j = 8, "H", "I") & ": Invalid price" & vbCrLf
                    End If
                Case 10  ' Order_Date (Column J)
                    If Not IsDate(cellValue) Then
                        errorCount = errorCount + 1
                        errorList = errorList & "Row " & i & ", Column J: Invalid date" & vbCrLf
                    End If
            End Select
        Next j
    Next i

    ' Display validation results
    If errorCount = 0 Then
        MsgBox "Data validation passed! No errors found.", vbInformation, "Validation Complete"
    Else
        MsgBox "Found " & errorCount & " errors:" & vbCrLf & vbCrLf & Left(errorList, 500) & _
               IIf(Len(errorList) > 500, "...", ""), vbExclamation, "Validation Errors"
    End If
End Sub
```

### Advanced Looping: Find and Replace

```vb
Sub UpdateProductNames()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim productName As String
    Dim updatedCount As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    updatedCount = 0

    ' Loop through product names and standardize them
    For i = 2 To lastRow
        productName = ws.Cells(i, 4).Value  ' Product_Name (Column D)

        ' Standardize product names
        Select Case True
            Case InStr(UCase(productName), "DRESS") > 0
                If UCase(productName) <> "SUMMER DRESS" And UCase(productName) <> "WINTER DRESS" Then
                    ws.Cells(i, 4).Value = "Fashion Dress"
                    updatedCount = updatedCount + 1
                End If
            Case InStr(UCase(productName), "SHOE") > 0 Or InStr(UCase(productName), "SNEAKER") > 0
                If UCase(productName) <> "RUNNING SHOES" And UCase(productName) <> "CASUAL SNEAKERS" Then
                    ws.Cells(i, 4).Value = "Fashion Footwear"
                    updatedCount = updatedCount + 1
                End If
            Case InStr(UCase(productName), "BAG") > 0
                If UCase(productName) <> "HANDBAG" And UCase(productName) <> "TOTE BAG" Then
                    ws.Cells(i, 4).Value = "Fashion Bag"
                    updatedCount = updatedCount + 1
                End If
        End Select
    Next i

    MsgBox "Product names standardized. " & updatedCount & " items updated.", vbInformation
End Sub
```

---

# 🔹 Session 4: Error Handling, Code Optimization, and Data Processing

## 4.1 Error Handling Techniques

### Understanding VBA Errors

Common VBA errors include:

- **Runtime Error 9**: Subscript out of range
- **Runtime Error 13**: Type mismatch
- **Runtime Error 1004**: Application-defined error
- **Runtime Error 91**: Object variable not set

### Basic Error Handling Structure

```vb
Sub BasicErrorHandling()
    On Error GoTo ErrorHandler

    ' Your main code here
    Dim ws As Worksheet
    Set ws = Worksheets("NonExistentSheet")  ' This will cause an error

    Exit Sub  ' Normal exit point

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error " & Err.Number
End Sub
```

### Advanced Error Handling for Fashion Data

```vb
Sub ProcessFashionOrdersWithErrorHandling()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim orderID As String
    Dim quantity As Variant
    Dim unitPrice As Variant
    Dim totalAmount As Double
    Dim successCount As Long
    Dim errorCount As Long
    Dim errorLog As String

    ' Initialize counters
    successCount = 0
    errorCount = 0
    errorLog = "Error Log:" & vbCrLf

    ' Set worksheet reference
    Set ws = Worksheets("Sales_Data")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Process each order
    For i = 2 To lastRow
        On Error Resume Next  ' Continue processing even if individual row has error

        orderID = ws.Cells(i, 1).Value
        quantity = ws.Cells(i, 7).Value
        unitPrice = ws.Cells(i, 8).Value

        ' Validate data types
        If Not IsNumeric(quantity) Or Not IsNumeric(unitPrice) Then
            errorCount = errorCount + 1
            errorLog = errorLog & "Row " & i & " (" & orderID & "): Invalid numeric data" & vbCrLf
            GoTo NextIteration
        End If

        ' Calculate total
        totalAmount = CDbl(quantity) * CDbl(unitPrice)

        ' Update the worksheet
        ws.Cells(i, 9).Value = totalAmount
        ws.Cells(i, 12).Value = "Processed: " & Now()

        successCount = successCount + 1

NextIteration:
        ' Clear any errors for next iteration
        Err.Clear
    Next i

    ' Restore normal error handling
    On Error GoTo 0

    ' Display summary
    MsgBox "Processing Complete!" & vbCrLf & _
           "Successful: " & successCount & vbCrLf & _
           "Errors: " & errorCount & vbCrLf & vbCrLf & _
           IIf(errorCount > 0, Left(errorLog, 300), ""), vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Critical Error: " & Err.Description & vbCrLf & _
           "Error Number: " & Err.Number, vbCritical, "System Error"
    On Error GoTo 0  ' Reset error handling
End Sub
```

### Error Handling with File Operations

```vb
Sub SafeFileOperations()
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim wb As Workbook
    Dim fileExists As Boolean

    filePath = "C:\Fashion_Reports\Monthly_Sales.xlsx"

    ' Check if file exists
    fileExists = (Dir(filePath) <> "")

    If fileExists Then
        ' Try to open existing file
        Set wb = Workbooks.Open(filePath)
        MsgBox "File opened successfully!", vbInformation
    Else
        ' Create new file
        Set wb = Workbooks.Add
        wb.SaveAs filePath
        MsgBox "New file created: " & filePath, vbInformation
    End If

    Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 1004
            MsgBox "Could not access file. Please check:" & vbCrLf & _
                   "1. File path exists" & vbCrLf & _
                   "2. File is not open in another application" & vbCrLf & _
                   "3. You have write permissions", vbExclamation
        Case 76
            MsgBox "Path not found. Creating directory...", vbInformation
            ' Create directory (requires additional code)
        Case Else
            MsgBox "Unexpected error: " & Err.Description, vbCritical
    End Select

    On Error GoTo 0
End Sub
```

## 4.2 Code Optimization with With...End With

### Without With Statement (Inefficient)

```vb
Sub FormatWithoutWith()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Inefficient - Excel object accessed multiple times
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1").Interior.Color = RGB(0, 102, 204)
    ws.Range("A1").HorizontalAlignment = xlCenter
    ws.Range("A1").Value = "Fashion Sales Report"
End Sub
```

### With Statement (Efficient)

```vb
Sub FormatWithWith()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Efficient - Excel object accessed once
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .Value = "Fashion Sales Report"
    End With
End Sub
```

### Complex With Statement Example

```vb
Sub FormatFashionReportOptimized()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Optimize screen updating
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Format headers efficiently
    With ws.Range("A1:K1")
        .Font.Bold = True
        .Font.Size = 12
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 102, 204)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlMedium
    End With

    ' Format data area efficiently
    With ws.Range("A2:K" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(128, 128, 128)
    End With

    ' Format specific columns
    With ws.Range("G2:G" & lastRow)  ' Quantity
        .NumberFormat = "0"
        .HorizontalAlignment = xlCenter
    End With

    With ws.Range("H2:I" & lastRow)  ' Prices
        .NumberFormat = "$#,##0.00"
        .HorizontalAlignment = xlRight
    End With

    With ws.Range("J2:J" & lastRow)  ' Dates
        .NumberFormat = "mm/dd/yyyy"
        .HorizontalAlignment = xlCenter
    End With

    ' Auto-fit columns
    ws.Columns("A:K").AutoFit

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    MsgBox "Report formatted efficiently!", vbInformation
End Sub
```

### Nested With Statements

```vb
Sub NestedWithExample()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    With ws
        With .Range("A1:K1")
            .Font.Bold = True
            With .Interior
                .Color = RGB(0, 102, 204)
                .Pattern = xlSolid
            End With
            With .Borders
                .LineStyle = xlContinuous
                .Weight = xlMedium
                .Color = RGB(255, 255, 255)
            End With
        End With

        ' Add title row
        With .Range("A1")
            .Value = "Fashion Sales Dashboard"
            With .Font
                .Size = 16
                .Color = RGB(255, 255, 255)
                .Bold = True
            End With
        End With
    End With
End Sub
```

## 4.3 Copy/Paste Operations via VBA

### Basic Copy/Paste Operations

```vb
Sub BasicCopyPaste()
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet

    Set sourceWS = Worksheets("Sales_Data")
    Set targetWS = Worksheets("Reports")

    ' Copy range
    sourceWS.Range("A1:K10").Copy

    ' Paste to target
    targetWS.Range("A1").PasteSpecial xlPasteAll

    ' Clear clipboard
    Application.CutCopyMode = False
End Sub
```

### Advanced Copy Operations for Fashion Data

```vb
Sub CopyFashionDataSelectively()
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim category As String

    Set sourceWS = Worksheets("Sales_Data")
    Set targetWS = Worksheets("Dress_Sales")

    ' Clear target sheet
    targetWS.Cells.Clear

    ' Copy headers
    sourceWS.Range("A1:K1").Copy
    targetWS.Range("A1").PasteSpecial xlPasteAll

    lastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row
    targetRow = 2

    ' Copy only dress-related sales
    For i = 2 To lastRow
        category = UCase(sourceWS.Cells(i, 3).Value)  ' Product_Category

        If InStr(category, "DRESS") > 0 Or InStr(category, "TOP") > 0 Then
            ' Copy entire row
            sourceWS.Range("A" & i & ":K" & i).Copy
            targetWS.Range("A" & targetRow).PasteSpecial xlPasteAll
            targetRow = targetRow + 1
        End If
    Next i

    Application.CutCopyMode = False
    MsgBox "Dress sales data copied successfully!", vbInformation
End Sub
```

### Copy with Transformation

```vb
Sub CopyAndTransformData()
    Dim sourceWS As Worksheet
    Dim summaryWS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim summaryRow As Long

    Set sourceWS = Worksheets("Sales_Data")

    ' Create or clear summary sheet
    On Error Resume Next
    Set summaryWS = Worksheets("Summary")
    On Error GoTo 0

    If summaryWS Is Nothing Then
        Set summaryWS = Worksheets.Add
        summaryWS.Name = "Summary"
    Else
        summaryWS.Cells.Clear
    End If

    ' Create summary headers
    With summaryWS
        .Range("A1").Value = "Customer Name"
        .Range("B1").Value = "Total Orders"
        .Range("C1").Value = "Total Amount"
        .Range("D1").Value = "Average Order"
        .Range("A1:D1").Font.Bold = True
    End With

    ' Get unique customers and their totals
    Dim customers As Object
    Set customers = CreateObject("Scripting.Dictionary")

    lastRow = sourceWS.Cells(sourceWS.Rows.Count, "A").End(xlUp).Row

    ' Collect customer data
    For i = 2 To lastRow
        Dim customerName As String
        Dim orderAmount As Double

        customerName = sourceWS.Cells(i, 2).Value
        orderAmount = sourceWS.Cells(i, 9).Value

        If customers.Exists(customerName) Then
            ' Update existing customer
            Dim existingData As Variant
            existingData = customers(customerName)
            existingData(0) = existingData(0) + 1  ' Order count
            existingData(1) = existingData(1) + orderAmount  ' Total amount
            customers(customerName) = existingData
        Else
            ' Add new customer
            customers(customerName) = Array(1, orderAmount)
        End If
    Next i

    ' Write summary to worksheet
    summaryRow = 2
    Dim customer As Variant
    For Each customer In customers.Keys
        Dim customerData As Variant
        customerData = customers(customer)

        With summaryWS
            .Cells(summaryRow, 1).Value = customer
            .Cells(summaryRow, 2).Value = customerData(0)  ' Order count
            .Cells(summaryRow, 3).Value = customerData(1)  ' Total amount
            .Cells(summaryRow, 4).Value = customerData(1) / customerData(0)  ' Average
        End With

        summaryRow = summaryRow + 1
    Next customer

    ' Format summary
    With summaryWS.Range("C2:D" & summaryRow - 1)
        .NumberFormat = "$#,##0.00"
    End With

    summaryWS.Columns("A:D").AutoFit

    MsgBox "Customer summary created with " & customers.Count & " customers!", vbInformation
End Sub
```

## 4.4 Filtering and Sorting Data

### Basic Sorting

```vb
Sub SortFashionData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set sortRange = ws.Range("A1:K" & lastRow)

    ' Sort by Total Amount (Column I) in descending order
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("I1:I" & lastRow), _
                       SortOn:=xlSortOnValues, _
                       Order:=xlDescending, _
                       DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    MsgBox "Data sorted by Total Amount (highest first)!", vbInformation
End Sub
```

### Multi-Level Sorting

```vb
Sub MultiLevelSort()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set sortRange = ws.Range("A1:K" & lastRow)

    ' Sort by Category (ascending), then by Total Amount (descending)
    With ws.Sort
        .SortFields.Clear
        ' Primary sort: Product Category
        .SortFields.Add Key:=ws.Range("C1:C" & lastRow), _
                       SortOn:=xlSortOnValues, _
                       Order:=xlAscending, _
                       DataOption:=xlSortNormal
        ' Secondary sort: Total Amount
        .SortFields.Add Key:=ws.Range("I1:I" & lastRow), _
                       SortOn:=xlSortOnValues, _
                       Order:=xlDescending, _
                       DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    MsgBox "Data sorted by Category, then by Amount!", vbInformation
End Sub
```

### AutoFilter Operations

```vb
Sub FilterFashionData()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim filterRange As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = 11  ' Columns A to K
    Set filterRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    ' Remove existing filters
    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' Apply AutoFilter
    filterRange.AutoFilter

    ' Filter for dresses only (assuming Category is in Column C)
    ws.AutoFilter.Range.AutoFilter Field:=3, Criteria1:="Dresses"

    MsgBox "Filtered to show Dresses only!", vbInformation
End Sub
```

### Advanced Filtering with Multiple Criteria

```vb
Sub AdvancedFilterExample()
    Dim ws As Worksheet
    Dim criteriaWS As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim criteriaRange As Range
    Dim outputRange As Range

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dataRange = ws.Range("A1:K" & lastRow)

    ' Create criteria sheet
    On Error Resume Next
    Set criteriaWS = Worksheets("Criteria")
    On Error GoTo 0

    If criteriaWS Is Nothing Then
        Set criteriaWS = Worksheets.Add
        criteriaWS.Name = "Criteria"
    Else
        criteriaWS.Cells.Clear
    End If

    ' Set up criteria
    With criteriaWS
        .Range("A1").Value = "Product_Category"
        .Range("B1").Value = "Total_Amount"
        .Range("A2").Value = "Dresses"
        .Range("B2").Value = ">100"
        .Range("A3").Value = "Shoes"
        .Range("B3").Value = ">75"
    End With

    Set criteriaRange = criteriaWS.Range("A1:B3")
    Set outputRange = ws.Range("M1")  ' Output starting at column M

    ' Apply advanced filter
    dataRange.AdvancedFilter Action:=xlFilterCopy, _
                            CriteriaRange:=criteriaRange, _
                            CopyToRange:=outputRange, _
                            Unique:=False

    MsgBox "Advanced filter applied! Results in columns M onwards.", vbInformation
End Sub
```

### Filter and Copy High-Value Orders

```vb
Sub FilterHighValueOrders()
    Dim ws As Worksheet
    Dim reportWS As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim reportRow As Long
    Dim totalAmount As Double
    Dim threshold As Double

    Set ws = Worksheets("Sales_Data")
    threshold = 100  ' Orders above $100

    ' Create report sheet
    On Error Resume Next
    Set reportWS = Worksheets("High_Value_Orders")
    On Error GoTo 0

    If reportWS Is Nothing Then
        Set reportWS = Worksheets.Add
        reportWS.Name = "High_Value_Orders"
    Else
        reportWS.Cells.Clear
    End If

    ' Copy headers
    ws.Range("A1:K1").Copy
    reportWS.Range("A1").PasteSpecial xlPasteAll

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    reportRow = 2

    ' Filter and copy high-value orders
    For i = 2 To lastRow
        totalAmount = ws.Cells(i, 9).Value  ' Total_Amount column

        If totalAmount > threshold Then
            ws.Range("A" & i & ":K" & i).Copy
            reportWS.Range("A" & reportRow).PasteSpecial xlPasteAll
            reportRow = reportRow + 1
        End If
    Next i

    Application.CutCopyMode = False

    ' Format the report
    With reportWS
        .Range("A1:K1").Font.Bold = True
        .Range("A1:K1").Interior.Color = RGB(255, 255, 0)
        .Columns("A:K").AutoFit
        .Range("H2:I" & reportRow - 1).NumberFormat = "$#,##0.00"
    End With

    MsgBox "High-value orders (>$" & threshold & ") report created!" & vbCrLf & _
           "Found " & (reportRow - 2) & " orders.", vbInformation
End Sub
```

## 4.5 Complete Fashion Sales Analysis Project

```vb
Sub CompleteFashionAnalysis()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim reportWS As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Performance optimization
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Create comprehensive report
    On Error Resume Next
    Set reportWS = Worksheets("Fashion_Analysis")
    On Error GoTo 0

    If reportWS Is Nothing Then
        Set reportWS = Worksheets.Add
        reportWS.Name = "Fashion_Analysis"
    Else
        reportWS.Cells.Clear
    End If

    ' Analysis variables
    Dim totalSales As Double
    Dim totalOrders As Long
    Dim avgOrderValue As Double
    Dim categoryStats As Object
    Set categoryStats = CreateObject("Scripting.Dictionary")

    ' Collect statistics
    For i = 2 To lastRow
        Dim category As String
        Dim orderAmount As Double

        category = ws.Cells(i, 3).Value
        orderAmount = ws.Cells(i, 9).Value

        totalSales = totalSales + orderAmount
        totalOrders = totalOrders + 1

        ' Category statistics
        If categoryStats.Exists(category) Then
            Dim catData As Variant
            catData = categoryStats(category)
            catData(0) = catData(0) + 1  ' Count
            catData(1) = catData(1) + orderAmount  ' Total
            categoryStats(category) = catData
        Else
            categoryStats(category) = Array(1, orderAmount)
        End If
    Next i

    avgOrderValue = totalSales / totalOrders

    ' Create report
    With reportWS
        ' Title
        .Range("A1").Value = "FASHION SALES ANALYSIS REPORT"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True

        ' Summary statistics
        .Range("A3").Value = "SUMMARY STATISTICS"
        .Range("A3").Font.Bold = True
        .Range("A4").Value = "Total Sales:"
        .Range("B4").Value = totalSales
        .Range("B4").NumberFormat = "$#,##0.00"
        .Range("A5").Value = "Total Orders:"
        .Range("B5").Value = totalOrders
        .Range("A6").Value = "Average Order Value:"
        .Range("B6").Value = avgOrderValue
        .Range("B6").NumberFormat = "$#,##0.00"

        ' Category breakdown
        .Range("A8").Value = "CATEGORY BREAKDOWN"
        .Range("A8").Font.Bold = True
        .Range("A9").Value = "Category"
        .Range("B9").Value = "Orders"
        .Range("C9").Value = "Total Sales"
        .Range("D9").Value = "Average"
        .Range("A9:D9").Font.Bold = True

        Dim reportRow As Long
        reportRow = 10

        Dim cat As Variant
        For Each cat In categoryStats.Keys
            Dim catInfo As Variant
            catInfo = categoryStats(cat)

            .Cells(reportRow, 1).Value = cat
            .Cells(reportRow, 2).Value = catInfo(0)
            .Cells(reportRow, 3).Value = catInfo(1)
            .Cells(reportRow, 3).NumberFormat = "$#,##0.00"
            .Cells(reportRow, 4).Value = catInfo(1) / catInfo(0)
            .Cells(reportRow, 4).NumberFormat = "$#,##0.00"

            reportRow = reportRow + 1
        Next cat

        ' Auto-fit columns
        .Columns("A:D").AutoFit

        ' Add borders and formatting
        .Range("A9:D" & reportRow - 1).Borders.LineStyle = xlContinuous
        .Range("A1:D" & reportRow - 1).Interior.ColorIndex = xlNone
        .Range("A9:D9").Interior.Color = RGB(200, 200, 200)
    End With

    ' Restore settings
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Complete fashion analysis report generated!" & vbCrLf & _
           "Total Sales: $" & Format(totalSales, "#,##0.00") & vbCrLf & _
           "Total Orders: " & totalOrders & vbCrLf & _
           "Categories: " & categoryStats.Count, vbInformation, "Analysis Complete"

    Exit Sub

ErrorHandler:
    ' Restore settings even if error occurs
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    MsgBox "Error in analysis: " & Err.Description, vbCritical, "Analysis Error"
End Sub
```

---

# 🎯 Practice Exercises

## Exercise 1: Basic VBA Skills

Create a macro that:

1. Adds sample fashion data to a worksheet
2. Formats the headers with colors and bold text
3. Calculates totals for each order
4. Displays a summary message

```vb
Sub Exercise1_Solution()
    ' Student implementation here
    ' Hint: Use variables, basic formatting, and message boxes
End Sub
```

## Exercise 2: Control Flow

Create a macro that:

1. Loops through fashion data
2. Categorizes orders as "Small" (<$50), "Medium" ($50-$150), or "Large" (>$150)
3. Counts orders in each category
4. Uses InputBox to get a minimum order value filter

```vb
Sub Exercise2_Solution()
    ' Student implementation here
    ' Hint: Use For loops, If statements, and InputBox
End Sub
```

## Exercise 3: Data Processing

Create a macro that:

1. Sorts data by customer name
2. Filters orders from the last 30 days
3. Creates a summary report by sales representative
4. Handles errors gracefully

```vb
Sub Exercise3_Solution()
    ' Student implementation here
    ' Hint: Use With statements, error handling, and sorting methods
End Sub
```

---

# 🔧 Best Practices and Tips

## Performance Optimization

1. **Turn off screen updating**: `Application.ScreenUpdating = False`
2. **Disable automatic calculation**: `Application.Calculation = xlCalculationManual`
3. **Use With statements** for multiple operations on the same object
4. **Declare variables with specific data types**
5. **Avoid selecting ranges** unless necessary

## Code Organization

1. **Use meaningful variable names**: `customerName` instead of `x`
2. **Add comments** to explain complex logic
3. **Break large procedures** into smaller, focused functions
4. **Use consistent indentation** for readability

## Error Prevention

1. **Always include error handling** in production code
2. **Validate user input** before processing
3. **Check if objects exist** before using them
4. **Use Option Explicit** at the top of modules

## Memory Management

1. **Set object variables to Nothing** when done: `Set ws = Nothing`
2. **Clear clipboard** after copy operations: `Application.CutCopyMode = False`
3. **Restore Excel settings** after changing them

---

# 📚 Additional Resources

## VBA Functions for Fashion Data

- **Text Functions**: `Left()`, `Right()`, `Mid()`, `InStr()`, `UCase()`, `LCase()`
- **Date Functions**: `Date()`, `Now()`, `DateAdd()`, `DateDiff()`, `Format()`
- **Math Functions**: `Round()`, `Sum()`, `Average()`, `Max()`, `Min()`
- **Validation**: `IsNumeric()`, `IsDate()`, `IsEmpty()`

## Excel Object Properties

- **Range.Value**: Get/set cell values
- **Range.Formula**: Get/set cell formulas
- **Range.NumberFormat**: Format numbers and dates
- **Range.Interior.Color**: Cell background color
- **Range.Font**: Font properties (Bold, Size, Color)

## Keyboard Shortcuts for VBA Development

- **Alt + F11**: Open/close VBA Editor
- **F5**: Run current procedure
- **F8**: Step through code line by line
- **Ctrl + G**: Open Immediate Window
- **Ctrl + R**: Show/hide Project Explorer

---

# 🎉 Congratulations!

You have completed the comprehensive VBA 101 tutorial! You now have the skills to:

✅ **Create and edit macros** to automate Excel tasks
✅ **Use variables and operators** for data manipulation
✅ **Implement control flow** with conditions and loops  
✅ **Handle user interaction** with message and input boxes
✅ **Work with Excel objects** (workbooks, worksheets, ranges)
✅ **Process data efficiently** with loops and filters
✅ **Handle errors gracefully** in your code
✅ **Optimize performance** using best practices
✅ **Sort and filter data** programmatically
✅ **Create comprehensive reports** and analysis tools

## Next Steps

1. **Practice regularly** with your own datasets
2. **Explore advanced topics** like UserForms and custom functions
3. **Learn about external data connections** and APIs
4. **Study other Office applications** (Word, PowerPoint VBA)
5. **Join VBA communities** for continued learning

Remember: The key to mastering VBA is **consistent practice** and **gradual progression** from simple to complex projects. Start with small automation tasks and gradually build more sophisticated solutions!

---

_Happy Coding! 🚀_
