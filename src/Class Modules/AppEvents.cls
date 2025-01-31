VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
'
' App_WorkbookOpen
'
' On existing workbook opened register udf help
'
' @author Matt Schuelke
' @copyright 2025
'

    ' Avoid: Run-time error ‘1004’ Cannot edit a macro on a hidden workbook. Unhide the workbook using the unhide command.
    If Workbooks.Count = 0 Then
        Dim tmpwb As Workbook
        Set tmpwb = Workbooks.Add
        I2DB.Registration.DoRegisterUDFs
        tmpwb.Close (False)
    Else
        I2DB.Registration.DoRegisterUDFs
    End If
    
End Sub

'Private Sub App_NewWorkbook(ByVal Wb As Workbook)
'
' App_NewWorkbook
'
' On new workbook created register udf help
'
' @author Matt Schuelke
' @copyright 2025
'
'
'    I2DB.Registration.DoRegisterUDFs
'
'End Sub

Private Sub Class_Terminate()
'
' Class_Terminate
'
' On class terminate cleanup
'
' @author Matt Schuelke
' @copyright 2025
'

    Set App = Nothing
    
End Sub

