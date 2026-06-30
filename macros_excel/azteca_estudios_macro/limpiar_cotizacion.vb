' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub LimpiarCotizacion()

    Dim i As Long
    Dim ultimaFila As Long
    Dim texto As String
    Dim filaInicioBloques As Long

    filaInicioBloques = 9

    ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row

    For i = filaInicioBloques To ultimaFila
        
        texto = Cells(i, "A").Value
        
        If InStr(1, texto, "IVA") > 0 Or InStr(1, texto, "SUBTOTAL") > 0 Then
            Exit For
        End If
        
        If InStr(1, texto, "SUMA TOTAL:") = 0 Then

            Cells(i, "A").ClearContents
            Cells(i, "G").ClearContents
            Cells(i, "H").ClearContents
            Cells(i, "I").ClearContents
            Cells(i, "J").ClearContents
            Cells(i, "L").ClearContents
            
        End If
        
    Next i

    MsgBox "La cotización se limpió correctamente ?", vbInformation

End Sub
