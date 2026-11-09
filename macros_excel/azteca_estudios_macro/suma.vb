' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Fernando Dorantes Nieto
' Modificado para mayor tolerancia y robustez
' 2026-11-09: trabaja SIEMPRE sobre la hoja Cotización (antes usaba la hoja
'             activa y podía escribir totales en otra hoja).
' 2026-11-09: rótulos por coincidencia EXACTA/prefijo (NormalizarRotulo y
'             EtiquetaTotal de M_Blindaje). Antes bastaba que un concepto
'             contuviera "SUMA" (p. ej. "Consumables") o una etiqueta "IVA"
'             (p. ej. "Activación") para cortar bloques o pisar totales.
' Depende de: HojaCotizacion(), FILA_INICIO (M_UUID)
'             NormalizarRotulo(), EtiquetaTotal() (M_Blindaje)
Sub suma()

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim inicioBloque As Long
    Dim ultimaFila As Long
    Dim ultimaI As Long
    Dim totalBloque As Double
    Dim subtotalGeneral As Double
    Dim valorIVA As Double
    Dim etiqueta As String
    Dim textoA As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ultimaI = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If ultimaI < ultimaFila Then ultimaI = ultimaFila
    subtotalGeneral = 0

    ' --- 1) Totales por bloque: cada "SUMA TOTAL" suma la columna M desde la
    '        SUMA TOTAL anterior (o desde FILA_INICIO si es el primer bloque).
    For i = FILA_INICIO To ultimaFila

        If NormalizarRotulo(ws.Cells(i, "A").Value) = "SUMATOTAL" Then

            totalBloque = 0
            inicioBloque = FILA_INICIO

            For j = i - 1 To FILA_INICIO Step -1
                If NormalizarRotulo(ws.Cells(j, "A").Value) = "SUMATOTAL" Then
                    inicioBloque = j + 1
                    Exit For
                End If
            Next j

            For j = inicioBloque To i - 1
                textoA = NormalizarRotulo(ws.Cells(j, "A").Value)
                If Len(textoA) > 0 And Left$(textoA, 7) <> "DESCRIP" And Left$(textoA, 6) <> "PRECIO" Then
                    If Not IsError(ws.Cells(j, "M").Value) Then
                        If IsNumeric(ws.Cells(j, "M").Value) And Len(CStr(ws.Cells(j, "M").Value)) > 0 Then
                            totalBloque = totalBloque + CDbl(ws.Cells(j, "M").Value)
                        End If
                    End If
                End If
            Next j

            ws.Cells(i, "M").Value = totalBloque
            subtotalGeneral = subtotalGeneral + totalBloque

        End If

    Next i

    ' --- 2) Totales generales (columna I = rótulo, columna K = valor) ---
    valorIVA = 0
    For i = 1 To ultimaI
        etiqueta = EtiquetaTotal(ws.Cells(i, "I").Value)
        Select Case etiqueta
            Case "SUBTOTAL"
                ws.Cells(i, "K").Value = subtotalGeneral
            Case "IVA"
                valorIVA = Round(subtotalGeneral * 0.16, 2)
                ws.Cells(i, "K").Value = valorIVA
            Case "TOTAL"
                ws.Cells(i, "K").Value = subtotalGeneral + valorIVA
        End Select
    Next i

End Sub
