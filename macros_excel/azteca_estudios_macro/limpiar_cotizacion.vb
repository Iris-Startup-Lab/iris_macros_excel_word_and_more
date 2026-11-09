' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Fernando Dorantes Nieto
' Modificado para confirmar acción, proteger títulos de categoría y recalcular a $0.00
' 2026-11-09: ya no usa ActiveSheet como respaldo (podía vaciar otra hoja) y no
'             intenta limpiar parcialmente celdas combinadas (error 1004).
Sub LimpiarCotizacion()

    Dim respuesta As VbMsgBoxResult
    Dim ws As Worksheet
    Dim i As Long
    Dim ultimaFila As Long
    Dim etiquetaA As String
    Dim pantalla As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización. No se limpió nada.", vbCritical, "Limpiar cotización"
        Exit Sub
    End If

    respuesta = MsgBox("¿Estás seguro de que deseas limpiar los datos de esta cotización?" & vbCrLf & vbCrLf & _
                       "Esta acción vaciará los conceptos capturados y no se puede deshacer.", _
                       vbYesNo + vbQuestion, "Confirmar limpieza")
    If respuesta <> vbYes Then Exit Sub

    FijarBotones ws

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > ultimaFila Then
        ultimaFila = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    End If

    pantalla = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo fallo

    For i = FILA_INICIO To ultimaFila

        ' Si llegamos al área de totales finales, detenemos la limpieza.
        ' 2026-11-09: por prefijo exacto; antes un concepto como "Activación"
        ' (contiene "IVA") detenía la limpieza a media cotización.
        etiquetaA = EtiquetaTotal(ws.Cells(i, "A").Value)
        If etiquetaA = "IVA" Or etiquetaA = "SUBTOTAL" Then Exit For
        If EtiquetaTotal(ws.Cells(i, "I").Value) = "SUBTOTAL" Then Exit For

        ' No tocar filas de suma, títulos de categoría combinados ni encabezados
        If Not EsFilaEstructural(ws, i) Then
            LimpiarCapturaFila ws, i
            ws.Cells(i, COL_MARCA).ClearContents
        End If

    Next i

    ' Recalcular para que subtotales y totales generales queden en $0.00
    Call suma

    RestaurarPosicionesBotones
    Application.ScreenUpdating = pantalla

    MsgBox "La cotización se limpió correctamente.", vbInformation, "Limpieza completada"
    Exit Sub

fallo:
    Application.ScreenUpdating = pantalla
    MsgBox "La limpieza se detuvo en la fila " & i & ":" & vbCrLf & Err.Description, _
           vbCritical, "Limpiar cotización"
End Sub

' Vacía B:L de una fila. Las celdas combinadas solo se limpian si la
' combinación completa cae dentro de B:L; si se sale (p. ej. un título A:N),
' se respeta, porque limpiar media combinación provoca el error 1004.
Private Sub LimpiarCapturaFila(ByVal ws As Worksheet, ByVal fila As Long)
    Dim col As Long
    Dim c As Range, m As Range

    For col = 2 To 12   ' B..L
        Set c = ws.Cells(fila, col)
        If c.MergeCells Then
            Set m = c.MergeArea
            If m.Cells(1, 1).Address = c.Address And _
               m.Column >= 2 And m.Column + m.Columns.Count - 1 <= 12 Then
                m.ClearContents
            End If
        Else
            c.ClearContents
        End If
    Next col
End Sub
