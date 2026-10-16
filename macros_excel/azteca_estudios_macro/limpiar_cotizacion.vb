' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado para confirmar acción, proteger títulos de categoría y recalcular a $0.00
Sub LimpiarCotizacion()

    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("¿Estás seguro de que deseas limpiar los datos de esta cotización?" & vbCrLf & vbCrLf & _
                       "Esta acción vaciará los conceptos capturados y no se puede deshacer.", _
                       vbYesNo + vbQuestion, "Confirmar limpieza")

    If respuesta <> vbYes Then Exit Sub

    FijarBotones

    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Set ws = ActiveSheet

    Dim i As Long
    Dim ultimaFila As Long
    Dim texto As String
    Dim textoNorm As String
    Dim filaInicioBloques As Long

    filaInicioBloques = FILA_INICIO

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > ultimaFila Then
        ultimaFila = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    End If

    Application.ScreenUpdating = False

    For i = filaInicioBloques To ultimaFila
        
        texto = Trim(CStr(ws.Cells(i, "A").Value))
        textoNorm = Replace(Replace(UCase(texto), ":", ""), " ", "")
        
        ' Si llegamos al área de totales finales, detenemos la limpieza
        If InStr(textoNorm, "IVA") > 0 Or InStr(textoNorm, "SUBTOTAL") > 0 Then
            Exit For
        End If
        
        ' Verificar si la columna I tiene SUBTOTAL
        If InStr(UCase(Trim(CStr(ws.Cells(i, "I").Value))), "SUBTOTAL") > 0 Then
            Exit For
        End If

        ' No tocar filas de suma ni encabezados
        If textoNorm <> "SUMATOTAL" Then
            ' No borrar si es encabezado de tabla (DESCRIPCION, PRECIO, etc.)
            If InStr(textoNorm, "DESCRIP") = 0 And InStr(textoNorm, "PRECIO") = 0 And _
               InStr(textoNorm, "CANTIDAD") = 0 And InStr(textoNorm, "ITEM") = 0 Then
               
                ' Limpiar únicamente columnas de captura de datos:
                ' B a L (descripción, unidades, cantidades, precios unitarios)
                ' dejando la columna A (categorías / estructura) intacta
                ws.Range(ws.Cells(i, "B"), ws.Cells(i, "L")).ClearContents
                
                ' Limpiar marca de detalle si la tenía
                ws.Cells(i, COL_MARCA).ClearContents
            End If
        End If
        
    Next i

    ' Recalcular para que subtotales y totales generales queden en $0.00
    Call suma

    RestaurarPosicionesBotones
    Application.ScreenUpdating = True

    MsgBox "La cotización se limpió correctamente.", vbInformation, "Limpieza completada"

End Sub
