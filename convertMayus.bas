Attribute VB_Name = "convertMayus"
Option Explicit
Public Function mayus(KeyAscii As Integer)
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase)) 'mayuscula
End Function

