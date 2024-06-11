Attribute VB_Name = "Registration"
Public Sub DoRegisterUDFs()
'
' DoRegisterUDFs
'
' Register the public UDFs defined in this add-in.
'
' @author Matt Schuelke
' @copyright 2024
'

    Hashing.RegisterUDFs
    
End Sub

Public Sub DeRegisterUDFs()
'
' DeRegisterUDFs
'
' DeRegister the public UDFs defined in this add-in.
'
' @author Matt Schuelke
' @copyright 2024
'

    Hashing.DeRegisterUDFs
    
End Sub

