Option explicit

' SHA256 usando la api de .NET.
' Funciona con datos binarios y retorna binarios.
' Tu mejor amigo será STRCONV.

Function BINSHA256(b() As Byte) As Byte()
    Dim objSHA256 As Object
    Set objSHA256 = CreateObject("System.Security.Cryptography.SHA256Managed")
    BINSHA256 = objSHA256.ComputeHash_2(b)
End Function
