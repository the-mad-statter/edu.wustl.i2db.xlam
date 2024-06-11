Attribute VB_Name = "Fill"
Public Sub EmptyCells(control As IRibbonControl)
'
' EmptyCells
'
' Fills selected empty cells with a user-specified string.
'
' @author Matt Schuelke
' @copyright 2017
'
'

    Dim message1 As String, message2 As String, message3 As String
    Dim r As String
    Dim c As Range
    
    On Error Resume Next
    
    message1 = "To use this function please ensure a range of cells is already highlighted."
    message2 = "Specify a string below which you would like entered into empty cells within your selection."
    message3 = "Press cancel to exit without changes."
    r = InputBox(message1 & vbCrLf & vbCrLf & message2 & vbCrLf & vbCrLf & message3, "Fill Empty Cells with String", "")
    
    For Each c In Selection
        If IsEmpty(c) Then
            c.value = r
        End If
    Next

End Sub
