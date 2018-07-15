Sub Macro_texto()
'
' Macro_texto para corregir el texto de los informes
'
'
Dim n_empresa, nombre_empresa As String
Dim colTexto As Collection

Set colTexto = New Collection

n_empresa = "ALICORP" 'Nombre corto de empresa
nombre_empresa = "ALICORP S.A.A." 'Razon social de la empresa

colTexto.Add "AGRÍCOLA DEL CHIRA S.A." 'Nombre largo de empresa anterior en informe
colTexto.Add nombre_empresa

colTexto.Add "AGRÍCOLA DEL CHIRA" 'Nombre corto de empresa anterior en informe
colTexto.Add n_empresa

colTexto.Add "n_empresa"
colTexto.Add n_empresa

colTexto.Add "N_empresa"
colTexto.Add n_empresa

colTexto.Add "nombre_empresa"
colTexto.Add nombre_empresa

colTexto.Add "Nombre_empresa"
colTexto.Add nombre_empresa

colTexto.Add " ,"
colTexto.Add ","

colTexto.Add " %"
colTexto.Add "%"

colTexto.Add "( "
colTexto.Add "("

colTexto.Add " )"
colTexto.Add ")"

colTexto.Add "formato F"
colTexto.Add "Formato F"

colTexto.Add "OFICIO - Nro."
colTexto.Add "Oficio N°"

colTexto.Add "CARTA - Nro."
colTexto.Add "Carta N°"

colTexto.Add "Carta N° SN"
colTexto.Add "Carta S/N"

colTexto.Add "carta N"
colTexto.Add "Carta N"

colTexto.Add "P-489"
colTexto.Add "Procedimiento"

Selection.Find.ClearFormatting
Selection.Find.Replacement.ClearFormatting

For i = 1 To colTexto.Count / 2
     With Selection.Find
        .Text = colTexto(i * 2 - 1)
        .Replacement.Text = colTexto(i * 2)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
Next i

Set colTexto = Nothing
    
End Sub
