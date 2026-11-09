' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Fernando Dorantes Nieto
' Modificado para mayor tolerancia y robustez
' 2026-11-09: trabaja SIEMPRE sobre la hoja Cotización (antes usaba la hoja
'             activa y podía escribir totales en otra hoja).
Sub suma()

    Dim ws As Worksheet
    Dim i As Long
    Dim j As Long
    Dim inicioBloque As Long
    Dim ultimaFila As Long
    Dim totalBloque As Double
    Dim subtotalGeneral As Double
    Dim texto As String
    Dim textoNorm As String
    Dim celdaTexto As Range
    Dim valorIVA As Double
    Dim textoCelda As String
    Dim x As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    subtotalGeneral = 0

    For i = 1 To ultimaFila

        texto = Trim(CStr(ws.Cells(i, "A").Value))
        textoNorm = Replace(Replace(UCase(texto), ":", ""), " ", "")

        If textoNorm = "SUMATOTAL" Then

            totalBloque = 0
            inicioBloque = 1

            For j = i - 1 To 1 Step -1
                If InStr(1, UCase(ws.Cells(j, "A").Value), "SUMA") > 0 Then
                    inicioBloque = j + 1
                    Exit For
                End If
            Next j

            For j = inicioBloque To i - 1

                If InStr(1, UCase(ws.Cells(j, "A").Value), "DESCRIP") > 0 Then GoTo Continuar
                If InStr(1, UCase(ws.Cells(j, "A").Value), "PRECIO") > 0 Then GoTo Continuar

                If IsNumeric(ws.Cells(j, "M").Value) Then
                    If Len(Trim(ws.Cells(j, "A").Value)) > 0 Then
                        totalBloque = totalBloque + CDbl(ws.Cells(j, "M").Value)
                    End If
                End If

Continuar:
            Next j

            ws.Cells(i, "M").Value = totalBloque
            subtotalGeneral = subtotalGeneral + totalBloque

        End If

    Next i

    For Each celdaTexto In ws.Range("I1:I" & ultimaFila)

        textoCelda = Replace(Replace(UCase(Trim(CStr(celdaTexto.Value))), ".", ""), " ", "")

        If InStr(textoCelda, "SUBTOTAL") > 0 Then
            ws.Cells(celdaTexto.Row, "K").Value = subtotalGeneral
        End If

        If InStr(textoCelda, "IVA") > 0 Then
            ws.Cells(celdaTexto.Row, "K").Value = Round(subtotalGeneral * 0.16, 2)
        End If

        If InStr(textoCelda, "TOTAL") > 0 And InStr(textoCelda, "SUBTOTAL") = 0 Then

            valorIVA = 0

            For x = celdaTexto.Row - 1 To 1 Step -1
                If InStr(Replace(Replace(UCase(CStr(ws.Cells(x, "I").Value)), ".", ""), " ", ""), "IVA") > 0 Then
                    valorIVA = CDbl(ws.Cells(x, "K").Value)
                    Exit For
                End If
            Next x

            ws.Cells(celdaTexto.Row, "K").Value = subtotalGeneral + valorIVA

        End If

    Next celdaTexto

End Sub
