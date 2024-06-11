Attribute VB_Name = "OffsetDelete"
Public Sub Row(control As IRibbonControl)
'
' Row
'
' After a contiguous range of cells is highlighted, this will allow one to delete a row every custom number of rows.
'
' @author Matt Schuelke
' @copyright 2024
'
'

    Dim message1 As String
    Dim message2 As String
    Dim message3 As String
    Dim message4 As String
    Dim title As String
    Dim enteredValue As Variant
    Dim i_by As Integer
    message1 = "To use this function please ensure a contiguous range of cells is already highlighted."
    message2 = "Entering a number below will delete the first row from the highlighted cells and then a single row every the-entered-number from the highlighted cells."
    message3 = "This action cannot be undone."
    message4 = "Press cancel to exit without changes."
    title = "Offset Row Deleter"
    enteredValue = InputBox(message1 & vbCrLf & vbCrLf & message2 & vbCrLf & vbCrLf & message3 & vbCrLf & vbCrLf & message4, title, "")
    If Not IsNumeric(enteredValue) Then
        MsgBox enteredValue & " is not numeric.", vbCritical
        Exit Sub
    Else
        i_by = CInt(enteredValue)
    End If
    
    ' correction to start deleting at first row
    Dim firstRowCorrection As Integer
    
    ' a cummulative corrective factor
    Dim nDeletedRows As Integer
    nDeletedRows = 0
    
    Dim correctedRow As Integer
    
    For Each rw In Selection.Rows
        If nDeletedRows = 0 Then
            firstRowCorrection = rw.Row Mod i_by
        End If
      
        correctedRow = rw.Row - firstRowCorrection + nDeletedRows
      
        If correctedRow Mod i_by = 0 Then
            rw.Delete shift:=xlUp
            nDeletedRows = nDeletedRows + 1
        End If
    Next rw
    
End Sub

Public Sub Col(control As IRibbonControl)
'
' Col
'
' After a contiguous range of cells is highlighted, this will allow one to delete a column every custom number of columns.
'
' @author Matt Schuelke
' @copyright 2024
'

    Dim message1 As String
    Dim message2 As String
    Dim message3 As String
    Dim message4 As String
    Dim title As String
    Dim enteredValue As Variant
    Dim i_by As Integer
    message1 = "To use this function please ensure a contiguous range of columns is already highlighted."
    message2 = "Entering a number below will delete the first column from the highlighted cells and then a single column every the-entered-number from the highlighted cells."
    message3 = "This action cannot be undone."
    message4 = "Press cancel to exit without changes."
    title = "Offset Column Deleter"
    enteredValue = InputBox(message1 & vbCrLf & vbCrLf & message2 & vbCrLf & vbCrLf & message3 & vbCrLf & vbCrLf & message4, title, "")
    If Not IsNumeric(enteredValue) Then
        MsgBox enteredValue & " is not numeric.", vbCritical
        Exit Sub
    Else
        i_by = CInt(enteredValue)
    End If
    
    ' correction to start deleting at first column
    Dim firstColumnCorrection As Integer
    
    ' a cummulative corrective factor
    Dim nDeletedColumns As Integer
    nDeletedColumns = 0
    
    Dim correctedColumn As Integer
    
    For Each cl In Selection.Columns
        If nDeletedColumns = 0 Then
            firstColumnCorrection = cl.Column Mod i_by
        End If
      
        correctedColumn = cl.Column - firstColumnCorrection + nDeletedColumns
      
        If correctedColumn Mod i_by = 0 Then
            cl.Delete shift:=xlToLeft
            nDeletedColumns = nDeletedColumns + 1
        End If
    Next cl
    
End Sub
