VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegistrationHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub Workbook_Open()

    Set App = Application
    
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)

    DoRegisterUDFs
    
End Sub
