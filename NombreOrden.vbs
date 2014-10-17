Public Function NombreOrden(strNombre As String) As String
'Ordena los nombres y apellidos, para casos con 2 apellidos y 2 nombres

Dim FullName As Variant

If Len(strNombre) = 0 Then
    NombreOrden = ""
Else
    FullName = Split(strNombre, " ")
    If UBound(FullName) <> 3 Then
        'UBound devuelve el indice mayor del array
        NombreOrden = "¡No tiene 4 nombres! " & StrConv(strNombre, vbProperCase)
    Else
        NombreOrden = StrConv(FullName(2) & " " & FullName(3) & " " & FullName(0) & " " & FullName(1), vbProperCase)
    End If
End If

End Function
