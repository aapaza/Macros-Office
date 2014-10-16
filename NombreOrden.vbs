Public Function NombreOrden(strNombre As String) As String
'Ordena los nombres y apellidos, para casos con 2 apellidos y 2 nombres

Dim FullName As Variant

If strNombre = "" Then
    NombreOrden = ""
Else
    FullName = Split(strNombre, " ")
    NombreOrden = StrConv(FullName(2) & " " & FullName(3) & " " & FullName(0) & " " & FullName(1), vbProperCase)
End If

End Function
