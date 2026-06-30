' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub suma()

    Dim i As Long
    Dim j As Long
    Dim inicioBloque As Long
    Dim ultimaFila As Long
    Dim totalBloque As Double
    Dim subtotalGeneral As Double
    Dim texto As String
    Dim celdaTexto As Range
    Dim valorIVA As Double
    Dim textoCelda As String
    Dim x As Long

    ultimaFila = Cells(Rows.Count, "A").End(xlUp).Row
    subtotalGeneral = 0

    ' =============================
    ' ?? 1. SUMA POR BLOQUES
    ' =============================
    For i = 1 To ultimaFila

        texto = Trim(Cells(i, "A").Value)

        If texto = "SUMA TOTAL" Or texto = "SUMA TOTAL:" Then

            totalBloque = 0
            inicioBloque = 1

            For j = i - 1 To 1 Step -1
                If InStr(1, Cells(j, "A").Value, "SUMA") > 0 Then
                    inicioBloque = j + 1
                    Exit For
                End If
            Next j

            For j = inicioBloque To i - 1

                If InStr(1, Cells(j, "A").Value, "DESCRIP") > 0 Then GoTo Continuar
                If InStr(1, Cells(j, "A").Value, "PRECIO") > 0 Then GoTo Continuar

                If IsNumeric(Cells(j, "M").Value) Then
                    If Len(Trim(Cells(j, "A").Value)) > 0 Then
                        totalBloque = totalBloque + Cells(j, "M").Value
                    End If
                End If

Continuar:
            Next j

            Cells(i, "M").Value = totalBloque
            subtotalGeneral = subtotalGeneral + totalBloque

        End If

    Next i

    ' =============================
    ' ?? 2. SUBTOTAL / IVA / TOTAL
    ' =============================
    For Each celdaTexto In Range("I1:I" & ultimaFila)

        ' ?? LIMPIAR TEXTO (clave)
        textoCelda = UCase(Replace(Trim(celdaTexto.Value), ".", ""))

        ' ? SUBTOTAL
        If InStr(textoCelda, "SUBTOTAL") > 0 Then
            Cells(celdaTexto.Row, "K").Value = subtotalGeneral
        End If

        ' ? IVA (YA FUNCIONA BIEN)
        If InStr(textoCelda, "IVA") > 0 Then
            Cells(celdaTexto.Row, "K").Value = subtotalGeneral * 0.16
        End If

        ' ? TOTAL FINAL
        If InStr(textoCelda, "TOTAL") > 0 Then

            valorIVA = 0

            For x = celdaTexto.Row - 1 To 1 Step -1
                If InStr(UCase(Replace(Cells(x, "I").Value, ".", "")), "IVA") > 0 Then
                    valorIVA = Cells(x, "K").Value
                    Exit For
                End If
            Next x

            Cells(celdaTexto.Row, "K").Value = subtotalGeneral + valorIVA

        End If

    Next celdaTexto

End Sub
