' LockManager.bas
' Handles dynamic locking and unlocking of columns in the Data sheet based on configuration.
Option Explicit

' Locks specific columns based on the provided configuration
' @param ws - The worksheet where locking will be applied
' @param lockedColumns - Array of column headers that need to be locked
Public Sub LockColumns(ws As Worksheet, lockedColumns As Variant)
    Dim col As Long
    Dim header As Range
    Dim columnName As Variant
    
    ' Unprotect the sheet before making changes
    ws.Unprotect
    
    ' Loop through the headers row and lock specified columns
    For Each header In ws.Rows(HEADER_ROW).Cells
        columnName = header.value
        
        ' Lock column if itâ€™s specified in lockedColumns array
        If Not IsError(Application.Match(columnName, lockedColumns, 0)) Then
            col = header.column
            ws.Columns(col).Locked = True
        Else
            ws.Columns(col).Locked = False
        End If
    Next header
    
    ' Re-protect the sheet after locking the columns
    ws.Protect password:="", AllowFiltering:=True
End Sub

' Unlocks all columns on the worksheet
' @param ws - The worksheet where all columns will be unlocked
Public Sub UnlockAllColumns(ws As Worksheet)
    ' Unprotect the sheet before unlocking all columns
    ws.Unprotect
    
    ' Unlock all columns in the worksheet
    ws.Cells.Locked = False
    
    ' Re-protect the sheet after unlocking columns
    ws.Protect password:="", AllowFiltering:=True
End Sub

' Main function to dynamically lock and unlock columns based on a specified list
' @param ws - The worksheet where locking will be applied
' @param lockedFields - Array of field names to lock based on the column headers in HEADER_ROW
Public Sub ApplyLocking(ws As Worksheet, lockedFields As Variant)
    ' Lock columns specified in lockedFields
    Call LockColumns(ws, lockedFields)
End Sub
