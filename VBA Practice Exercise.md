## 📝 VBA Practice Exercise: "Sales Report File Reader with Error Handling"

### 🔧 Scenario:

You're creating a macro for your finance team that will **read monthly sales data** from a `.txt` file and display its contents in a message box. The file is located at a specific path, but sometimes:

* The file may not exist.
* The file may be open in another program.
* The file may be empty.

Your task is to implement proper **error handling** to manage all these situations gracefully.

---

### 💡 Requirements:

1. The macro should attempt to open the file `C:\Reports\Sales_August.txt`.
2. If the file is not found, show an error message:
   `"Error: File not found."`
3. If the file is empty, show an error message:
   `"Error: File is empty."`
4. If the file opens successfully, read the entire content and display it using `MsgBox`.
5. In case of any other unexpected error, show a message like:
   `"Unexpected Error: <error description>"`

---

### 📌 Starter Code:

```vba
Sub ReadSalesReport()
    On Error GoTo ErrorHandler

    Dim filePath As String
    Dim fileNum As Integer
    Dim fileContent As String
    Dim lineText As String

    filePath = "C:\Reports\Sales_August.txt"
    fileNum = FreeFile

    ' Attempt to open file
    Open filePath For Input As #fileNum

    ' Read file line by line
    Do Until EOF(fileNum)
        Line Input #fileNum, lineText
        fileContent = fileContent & lineText & vbCrLf
    Loop

    Close #fileNum

    ' Check for empty file
    If Trim(fileContent) = "" Then
        MsgBox "Error: File is empty."
        Exit Sub
    End If

    ' Show file content
    MsgBox "Sales Report Content:" & vbCrLf & fileContent
    Exit Sub

ErrorHandler:
    If Err.Number = 53 Then ' File not found
        MsgBox "Error: File not found."
    Else
        MsgBox "Unexpected Error: " & Err.Description
    End If
End Sub
```
