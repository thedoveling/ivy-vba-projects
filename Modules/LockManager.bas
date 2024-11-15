' LockManager.bas
Option Explicit

' Locks or unlocks columns based on configuration.
' @param targetSheet - The Excel sheet to lock/unlock columns in
' @param lockConfig - A dictionary of columns to lock/unlock
Public Sub ManageLocks(targetSheet As Worksheet, lockConfig As Scripting.Dictionary)
    Dim col As Variant
    For Each col In lockConfig.Keys
        targetSheet.Columns(col).Locked = lockConfig(col)
    Next col
    targetSheet.Protect
End Sub