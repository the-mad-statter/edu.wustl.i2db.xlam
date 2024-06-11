VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} About 
   Caption         =   "About edu.wustl.i2db.xlam"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "About.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Unload Me
End Sub

Private Sub Label3_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/the-mad-statter/edu.wustl.i2db.xlam", NewWindow:=True
End Sub

Private Sub Label4_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://github.com/the-mad-statter/edu.wustl.i2db.xlam/issues", NewWindow:=True
End Sub

Private Sub Label5_Click()
    ActiveWorkbook.FollowHyperlink Address:="https://the-mad-statter.github.io/edu.wustl.i2db.xlam", NewWindow:=True
End Sub
