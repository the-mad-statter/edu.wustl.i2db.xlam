Attribute VB_Name = "Trim"
Public Function I2DB_CODE(ByVal value As String)
'
' I2DB_CODE
'
' Vectorized version of CODE() that converts each character of a string from Unicode to the default code page of the system and spills results to the right
'
' @author Matt Schuelke
' @copyright 2025

    ' VBA version of `=CODE(MID(A1,SEQUENCE(1,LEN(A1)),1))`
    Dim bytes() As Byte
    bytes = StrConv(value, vbFromUnicode)
    I2DB_CODE = bytes
    
End Function

Public Function I2DB_TRIM(ByVal value As String)
'
' I2DB_TRIM
'
' Version of TRIM() that removes all nonprintable characters including non-breaking spaces
'
' @author Matt Schuelke
' @copyright 2025
'

    Dim wsf As Object
    Set wsf = Application.WorksheetFunction
    I2DB_TRIM = wsf.Trim(wsf.Clean(wsf.Substitute(value, Chr(160), "")))

End Function

Public Sub RegisterUDFs()
'
' RegisterUDFs
'
' Register the public UDFs defined in this module.
'
' @author Matt Schuelke
' @copyright 2025
'

    Dim sFunDescr As String
    sFunDescr = "Vectorized version of CODE() that converts each character of a string from Unicode to the default code page of the system and spills results to the right"
    Dim vArgDescr() As Variant
    ReDim vArgDescr(1)
    vArgDescr = "String to convert"
        
    Application.MacroOptions _
        Macro:="I2DB_CODE", _
        Description:=sFunDescr, _
        Category:="I2DB", _
        ArgumentDescriptions:=vArgDescr

    sFunDescr = "Version of TRIM() that removes all nonprintable characters including non-breaking spaces"
    vArgDescr = "String to trim"
        
    Application.MacroOptions _
        Macro:="I2DB_TRIM", _
        Description:=sFunDescr, _
        Category:="I2DB", _
        ArgumentDescriptions:=vArgDescr
        
End Sub

Public Sub DeRegisterUDFs()
'
' DeRegisterUDFs
'
' DeRegister the public UDFs defined in this module.
'
' @author Matt Schuelke
' @copyright 2025
'

    Dim sFunDescr As String
    Dim vArgDescr() As Variant
    ReDim vArgDescr(1)
    
    Application.MacroOptions _
        Macro:="I2DB_CODE", _
        Description:=sFunDescr, _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
    
    Application.MacroOptions _
        Macro:="I2DB_TRIM", _
        Description:=sFunDescr, _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
        
End Sub
