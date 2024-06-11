Attribute VB_Name = "Insert"
Public Sub MultiRows(control As IRibbonControl)
'
' MultiRows
'
' Insert a custom number of rows above the current selection.
'
' @author Matt Schuelke
' @copyright 2024
'

    s = InputBox("Enter number of rows to insert.")
    If Not IsNumeric(s) Then
        MsgBox s & " is not numeric.", vbCritical
        Exit Sub
    Else
        n = CInt(s)
    End If
    ActiveCell.EntireRow.Resize(n).Insert shift:=xlDown

End Sub

Public Sub MultiCols(control As IRibbonControl)
'
' MultiCols
'
' Insert a custom number of columns before the current selection.
'
' @author Matt Schuelke
' @copyright 2024
'

    s = InputBox("Enter number of columns to insert.")
    If Not IsNumeric(s) Then
        MsgBox s & " is not numeric.", vbCritical
        Exit Sub
    Else
        n = CInt(s)
    End If
    ActiveCell.EntireColumn.Resize(, n).Insert shift:=xlRight

End Sub

