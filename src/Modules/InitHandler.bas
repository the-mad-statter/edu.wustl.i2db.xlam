Attribute VB_Name = "InitHandler"
Option Explicit

Dim EventHandler As AppEvents

Sub InitHandler()
'
' InitHandler
'
' Create instance of AppEvents defined in Class Module AppEvents
'
' @author Matt Schuelke
' @copyright 2025
'

    Set EventHandler = New AppEvents
    Set EventHandler.App = Application
    
End Sub
