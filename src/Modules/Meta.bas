Attribute VB_Name = "Meta"
Public Sub ReportHiddenDims(control As IRibbonControl)
'
' ReportHiddenDims
'
' Report hidden rows and columns in a message box.
'
' @author Matt Schuelke
' @copyright 2024
'

    Dim r As Range, s As String, msg As String
    
    sRowsMax = InputBox("Enter maximum number of rows to examine.", "", 1000)
    If Not IsNumeric(sRowsMax) Then
        MsgBox sRowsMax & " is not numeric.", vbCritical
        Exit Sub
    Else
        nRowsMax = CInt(sRowsMax)
    End If
    For Each r In Range(Rows(1), Rows(nRowsMax))
        If r.Hidden Then
            msg = msg & "Row " & r.Row & vbNewLine
        End If
    Next r
    
    sColsMax = InputBox("Enter maximum number of columns to examine.", "", 100)
    If Not IsNumeric(sColsMax) Then
        MsgBox sColsMax & " is not numeric.", vbCritical
        Exit Sub
    Else
        nColsMax = CInt(sColsMax)
    End If
    
    For Each r In Range(Columns(1), Columns(nColsMax))
        If r.Hidden Then
            s = Split(r.Address(, False), ":")(0)
            msg = msg & "Col " & s & vbNewLine
        End If
    Next r
    
    If msg = "" Then
        msg = "There are no hidden rows/cols"
    Else
        s = "The following rows/cols are hidden:"
        msg = s & vbNewLine & vbNewLine & msg
    End If
    
    MsgBox msg, vbInformation, "Hidden Rows and Columns"

End Sub
