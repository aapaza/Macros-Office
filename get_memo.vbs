Sub getMemo()
'Macro para combinar datos de Excel en un archivo en Word
 
    'Dim MiHoja As Object
    Dim plantilla(2) As String
    Dim curCell As Range
    Set curCell = ActiveCell
   
    Dim strNombreEmpresaCorto, strOrigen(8), strDestino(8) As String
   
    plantilla(1) = "\\psf\Home\Downloads\UGSEIN Memo 2014 IT - Evaluacion Descargos PAS ERACG 2012 - Plantilla.doc"
    plantilla(2) = "\\psf\Home\Downloads\UGSEIN Memo 2014 IT - Evaluacion Descargos PAS ERACG 2012 - Plantilla rpta reiteracion.doc"
    'poner entre las comillas la ruta del documento y la extensión del archivo, en este caso documento de Word habilitado para macros.
    Dim objWord, objDoc As Object
      
    Set objWord = CreateObject("Word.Application")
    'objDoc.SaveAs ("e:\test\temp.docx")
    'objWord.Quit
    
    strNombreEmpresaCorto = Cells(curCell.Row, 5).Value
 
    strOrigen(1) = "$fecha$" 'fecha de hoy para el memo
    strOrigen(2) = "$nombre_empresa$"
    strOrigen(3) = "$numero_oficio$"
    strOrigen(4) = "$numero_expediente$"
    strOrigen(5) = "$codigo_IT$"
    strOrigen(6) = "$memo_alfe$"
    strOrigen(7) = "$numero_expediente_legal$"
'===
'En el 2013 no hay memo de reiteración
'    strOrigen(7) = "$memo_alfe_reitera$"
'    strOrigen(8) = "$expediente_legal$"
'===
'2013
    strDestino(1) = Format(Now(), "dd \de mmmm \de yyyy")
    strDestino(2) = Cells(curCell.Row, 2).Value 'nombre de empresa
    strDestino(3) = Replace(Cells(curCell.Row, 40).Value, "OFICIO - Nro. ", "") 'Número de Oficio
    strDestino(4) = Cells(curCell.Row, 12).Value 'Número de expediente
    strDestino(5) = Cells(curCell.Row, 51).Value 'Código del informe técnico
    strDestino(6) = Replace(Cells(curCell.Row, 44).Value, "MEMORANDUM - Nro. ", "") 'Número de memo ALFE con el que piden analizar el descargo
    strDestino(7) = Replace(Cells(curCell.Row, 45).Value, "N° ", "") 'Número de memo ALFE con el que piden analizar el descargo
'===
'En el 2013 no hay memo de reiteración
'    strDestino(7) = Cells(curCell.Row, 41).Value
'    strDestino(8) = Cells(curCell.Row, 36).Value
'===
'2012
'    strDestino(1) = Format(Now(), "dd \de mmmm \de yyyy")
'    strDestino(2) = Cells(curCell.Row, 5).Value
'    strDestino(3) = Replace(Cells(curCell.Row, 31).Value, "OFICIO - Nro. ", "")
'    strDestino(4) = Cells(curCell.Row, 12).Value
'    strDestino(5) = Cells(curCell.Row, 38).Value
'    strDestino(6) = Replace(Cells(curCell.Row, 35).Value, "MEMORANDUM - Nro. ", "")
'    strDestino(7) = Cells(curCell.Row, 41).Value
'    strDestino(8) = Cells(curCell.Row, 36).Value
        
    For k = 1 To 2
        Set objDoc = objWord.Documents.Add(plantilla(k))
        objWord.Visible = False
        
        For i = 1 To 7
            With objWord.Selection
                .Find.ClearFormatting
                .Move 6, -1 'moverse al principio del documento
                .Find.Text = strOrigen(i)
                .Find.Execute
                Do While .Find.Found
                    .Text = strDestino(i)
                    .Move 6, -1 'moverse al principio del documento
                    .Find.Execute
                Loop
            End With
        Next
        
        If k = 1 Then
            objDoc.SaveAs ("\\psf\Dropbox\1. ERACG ISO 9001\GFE-UGSE-PE-03\Registros del Proceso\Registros 2013\4. Informes Tecnicos de descargo\UGSEIN Memo 2015 IT - Evaluacion Descargos PAS ERACG 2013 - " & strNombreEmpresaCorto & ".doc")
            'Para el 2012 era este valor:
            'objDoc.SaveAs ("\\psf\Home\Downloads\UGSEIN Memo 2014 IT - Evaluacion Descargos PAS ERACG 2013 - " & strNombreEmpresaCorto & ".doc")
        Else
            objDoc.SaveAs ("\\psf\Dropbox\1. ERACG ISO 9001\GFE-UGSE-PE-03\Registros del Proceso\Registros 2013\4. Informes Tecnicos de descargo\UGSEIN Memo 2015 IT - Evaluacion Descargos PAS ERACG 2013 - " & strNombreEmpresaCorto & " rpta reiteracion.doc")
        End If
        objDoc.Close
    Next
    'Salir de word
    objWord.Quit
    Set objWord = Nothing
 
End Sub
