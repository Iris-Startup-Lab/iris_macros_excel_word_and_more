' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado con validación de contexto de fila y recálculo automático
' 2026-11-09: busca la fila "SUMA TOTAL" por coincidencia exacta (antes cualquier
'             concepto que contuviera "SUMA" se tomaba como fin de bloque), opera
'             sobre la hoja Cotización explícitamente y valida la hoja Plantillas.
Sub DuplicarBloque()

    Dim hojaDestino As Worksheet
    Dim hojaPlantilla As Worksheet
    Dim filaActual As Long
    Dim filaDestino As Long
    Dim ultimaFila As Long
    Dim r As Long
    Dim numeroFilasBloque As Long
    Dim filasExtraAbajo As Long

    filaActual = ActiveCell.Row
    If Not ValidarFilaEnBloque(filaActual) Then Exit Sub

    Set hojaDestino = HojaCotizacion()

    On Error Resume Next
    Set hojaPlantilla = ThisWorkbook.Worksheets("Plantillas")
    On Error GoTo 0
    If hojaPlantilla Is Nothing Then
        MsgBox "No se encontró la hoja «Plantillas» con el bloque modelo.", vbCritical, "Duplicar bloque"
        Exit Sub
    End If

    ultimaFila = hojaDestino.Cells(hojaDestino.Rows.Count, "A").End(xlUp).Row
    filaDestino = 0
    For r = filaActual To ultimaFila
        If NormalizarRotulo(hojaDestino.Cells(r, "A").Value) = "SUMATOTAL" Then
            filaDestino = r + 2
            Exit For
        End If
    Next r

    If filaDestino = 0 Then
        MsgBox "No se encontró ninguna fila de SUMA TOTAL debajo de la fila seleccionada.", vbExclamation, "No se encontró bloque"
        Exit Sub
    End If

    FijarBotones hojaDestino    ' insertar un bloque completo es lo que más los movía

    numeroFilasBloque = hojaPlantilla.Range("A7:N12").Rows.Count
    filasExtraAbajo = 2

    hojaDestino.Rows(filaDestino & ":" & filaDestino + numeroFilasBloque + filasExtraAbajo - 1).Insert Shift:=xlDown
    hojaPlantilla.Range("A7:N12").Copy
    hojaDestino.Range("A" & filaDestino).PasteSpecial xlPasteAll
    Application.CutCopyMode = False

    RestaurarPosicionesBotones
    Call suma               ' recalcular totales

End Sub
