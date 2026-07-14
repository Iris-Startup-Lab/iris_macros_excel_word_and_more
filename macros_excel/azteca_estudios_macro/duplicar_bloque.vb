' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub DuplicarBloque()

    Dim hojaDestino As Worksheet
    Dim hojaPlantilla As Worksheet
    Dim filaActual As Long
    Dim filaDestino As Long
    Dim celda As Range
    Dim numeroFilasBloque As Long
    Dim filasExtraAbajo As Long

    Set hojaDestino = ActiveSheet
    Set hojaPlantilla = Sheets("Plantillas")

    filaActual = ActiveCell.Row
    filaDestino = 0

    For Each celda In hojaDestino.Range("A" & filaActual & ":A1000")
        If InStr(1, celda.Value, "SUMA") > 0 Then
            filaDestino = celda.Row + 2
            Exit For
        End If
    Next celda

    If filaDestino = 0 Then
        MsgBox "No se encontró ninguna fila de SUMA debajo."
        Exit Sub
    End If

    numeroFilasBloque = hojaPlantilla.Range("A7:N12").Rows.Count

    filasExtraAbajo = 2

    Rows(filaDestino & ":" & filaDestino + numeroFilasBloque + filasExtraAbajo - 1).Insert Shift:=xlDown

    hojaPlantilla.Range("A7:N12").Copy

    hojaDestino.Range("A" & filaDestino).PasteSpecial xlPasteAll

    Application.CutCopyMode = False

End Sub