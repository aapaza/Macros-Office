Public Function NombreOrden(strNombre As String) As String
'Ordena los nombres y apellidos, considera exactamente 2 nombres y 2 apellidos

Dim FullName As Variant

FullName = Split(strNombre, " ")
NombreOrden = StrConv(FullName(2) & " " & FullName(3) & " " & FullName(0) & " " & FullName(1), vbProperCase)

End Function
