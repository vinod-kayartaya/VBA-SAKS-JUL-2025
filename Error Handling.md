# Error handling in VBA

Error handling in **VBA (Visual Basic for Applications)** is the process of anticipating, capturing, and responding to runtime errors — that is, errors that occur while your code is running (not during compilation). It helps prevent your program from crashing and allows you to provide meaningful error messages or alternate flows.

---

## ✅ Types of Errors in VBA

1. **Syntax Errors**: Mistakes in code structure (e.g., missing `End If`, misspelled keywords). Caught at compile time.
2. **Runtime Errors**: Occur during execution (e.g., divide by zero, file not found).
3. **Logical Errors**: Code runs but gives wrong result.

---

## ✅ Basic Error Handling Using `On Error`

### 1. `On Error Resume Next`

Skips the line where the error occurs and continues with the next line of code.

```vba
Sub Example1()
    On Error Resume Next
    Dim x As Integer
    x = 10 / 0  ' Causes division by zero
    MsgBox "Execution continues despite the error"
End Sub
```

🧠 **Use with caution:** It ignores all errors and may hide bugs.

---

### 2. `On Error GoTo Label`

Redirects execution to a labeled block when an error occurs.

```vba
Sub Example2()
    On Error GoTo ErrorHandler

    Dim x As Integer
    x = 10 / 0  ' This will cause an error

    MsgBox "This will not execute if an error occurs"
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub
```

---

### 3. `On Error GoTo 0`

Disables any active error handler. Errors will now trigger default error dialog.

```vba
Sub Example3()
    On Error GoTo ErrorHandler

    Dim x As Integer
    x = 10 / 0

    On Error GoTo 0  ' Error handling turned off
    Exit Sub

ErrorHandler:
    MsgBox "Error handled"
End Sub
```

---

## ✅ Understanding the `Err` Object

The `Err` object contains details about the error:

| Property          | Description                                         |
| ----------------- | --------------------------------------------------- |
| `Err.Number`      | Numeric code of the error                           |
| `Err.Description` | Text description of the error                       |
| `Err.Source`      | Name of the project or object that caused the error |
| `Err.Clear`       | Clears the error from memory                        |

Example:

```vba
MsgBox "Error " & Err.Number & ": " & Err.Description
```

---

## ✅ Example: Error Handling with File Access

```vba
Sub OpenFile()
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile
    Open "C:\nonexistentfile.txt" For Input As #fileNum

    Close #fileNum
    Exit Sub

ErrorHandler:
    MsgBox "File error: " & Err.Description
End Sub
```

---

## ✅ Best Practices

- Always **exit the subroutine/function** before the error handler using `Exit Sub` or `Exit Function`.
- Clean up resources (close files, release objects) in the error handler if necessary.
- Use `Err.Clear` to reset the error object if you handle the error and want to continue.
- Don’t overuse `On Error Resume Next`.

---

## ✅ Advanced: Using `Resume` Statements

| Statement      | Behavior                                            |
| -------------- | --------------------------------------------------- |
| `Resume`       | Re-executes the line that caused the error.         |
| `Resume Next`  | Executes the line immediately after the error line. |
| `Resume Label` | Jumps to a specific line label.                     |

---
