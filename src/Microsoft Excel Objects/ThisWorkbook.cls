VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
'
' Workbook_Open
'
' On add-in workbook open call InitHandler() in Module InitHandler
'
' @author Matt Schuelke
' @copyright 2025
'

    Call InitHandler.InitHandler
    
End Sub
