Attribute VB_Name = "Registration"
Public Sub DoRegisterUDFs()
'
' DoRegisterUDFs
'
' Register the public UDFs defined in this add-in.
'
' @author Matt Schuelke
' @copyright 2025
'

    Hashing.RegisterUDFs
    Trim.RegisterUDFs
    
End Sub

Public Sub DeRegisterUDFs()
'
' DeRegisterUDFs
'
' DeRegister the public UDFs defined in this add-in.
'
' @author Matt Schuelke
' @copyright 2025
'

    Hashing.DeRegisterUDFs
    Trim.DeRegisterUDFs
    
End Sub
