Attribute VB_Name = "Hashing"
Public Function I2DB_HMACSHA256(ByVal value As String, ByVal key As String)
Attribute I2DB_HMACSHA256.VB_Description = "Compute a Hash-based Message Authentication Code (HMAC) using the SHA256 hash function."
Attribute I2DB_HMACSHA256.VB_ProcData.VB_Invoke_Func = " \n19"
'
' I2DB_HMACSHA256
'
' Compute a Hash-based Message Authentication Code (HMAC) using the SHA256 hash function.
'
' @author Matt Schuelke
' @copyright 2024
'

    Dim asc As Object
    Dim enc As Object
    Dim bValue() As Byte
    Dim bKey() As Byte
    Dim b() As Byte

    Set asc = CreateObject("System.Text.UTF8Encoding")
    Set enc = CreateObject("System.Security.Cryptography.HMACSHA256")

    bValue = asc.GetBytes_4(value)
    bKey = asc.GetBytes_4(key)
    enc.key = bKey

    b = enc.ComputeHash_2((bValue))
    I2DB_HMACSHA256 = EncodeBase64(b)

    Set asc = Nothing
    Set enc = Nothing

End Function

Private Function EncodeBase64(ByRef arrData() As Byte) As String
'
' EncodeBase64
'
' Returns a base 64 string version of passed byte array
'
' @author anonymous
' @copyright 2012
'

    Dim objXML As Object
    Dim objNode As Object

    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")

    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    EncodeBase64 = objNode.Text

    Set objNode = Nothing
    Set objXML = Nothing

End Function

Public Sub RegisterUDFs()
'
' RegisterUDFs
'
' Register the public UDFs defined in this module.
'
' @author Matt Schuelke
' @copyright 2024
'

    Dim vArgDescr() As Variant
    ReDim vArgDescr(1 To 2)
    vArgDescr(1) = "Value to hash"
    vArgDescr(2) = "Secret key value"
        
    Application.MacroOptions _
        Macro:="I2DB_HMACSHA256", _
        Description:="Compute a Hash-based Message Authentication Code (HMAC) using the SHA256 hash function.", _
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
' @copyright 2024
'

    Dim sFunDescr As String
    Dim vArgDescr() As Variant
    ReDim vArgDescr(1)
    
    Application.MacroOptions _
        Macro:="I2DB_HMACSHA256", _
        Description:=sFunDescr, _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
        
End Sub

