' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado para mayor tolerancia y robustez
Sub suma()

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

    ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row
    subtotalGeneral = 0

    For i = 1 To ultimaFila

        texto = Trim(CStr(Cells(i, "A").Value))
        textoNorm = Replace(Replace(UCase(texto), ":", ""), " ", "")

        If textoNorm = "SUMATOTAL" Then

            totalBloque = 0
            inicioBloque = 1

            For j = i - 1 To 1 Step -1
                If InStr(1, UCase(Cells(j, "A").Value), "SUMA") > 0 Then
                    inicioBloque = j + 1
                    Exit For
                End If
            Next j

            For j = inicioBloque To i - 1

                If InStr(1, UCase(Cells(j, "A").Value), "DESCRIP") > 0 Then GoTo Continuar
                If InStr(1, UCase(Cells(j, "A").Value), "PRECIO") > 0 Then GoTo Continuar

                If IsNumeric(Cells(j, "M").Value) Then
                    If Len(Trim(Cells(j, "A").Value)) > 0 Then
                        totalBloque = totalBloque + CDbl(Cells(j, "M").Value)
                    End If
                End If

Continuar:
            Next j

            Cells(i, "M").Value = totalBloque
            subtotalGeneral = subtotalGeneral + totalBloque

        End If

    Next i

    For Each celdaTexto In Range("I1:I" & ultimaFila)

        textoCelda = Replace(Replace(UCase(Trim(CStr(celdaTexto.Value))), ".", ""), " ", "")

        If InStr(textoCelda, "SUBTOTAL") > 0 Then
            Cells(celdaTexto.Row, "K").Value = subtotalGeneral
        End If

        If InStr(textoCelda, "IVA") > 0 Then
            Cells(celdaTexto.Row, "K").Value = Round(subtotalGeneral * 0.16, 2)
        End If

        If InStr(textoCelda, "TOTAL") > 0 And InStr(textoCelda, "SUBTOTAL") = 0 Then

            valorIVA = 0

            For x = celdaTexto.Row - 1 To 1 Step -1
                If InStr(Replace(Replace(UCase(CStr(Cells(x, "I").Value)), ".", ""), " ", ""), "IVA") > 0 Then
                    valorIVA = CDbl(Cells(x, "K").Value)
                    Exit For
                End If
            Next x

            Cells(celdaTexto.Row, "K").Value = subtotalGeneral + valorIVA

        End If

    Next celdaTexto

End Sub
