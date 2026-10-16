' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado para validar contexto de celda activa, proteger encabezados y recalcular totales
Sub BorrarFilaSeleccionada()

    Dim fila As Long
    Dim texto As String
    Dim textoNorm As String
    Dim respuesta As VbMsgBoxResult

    fila = ActiveCell.Row
    
    ' 1) Validar que la fila esté dentro del bloque de conceptos de la Cotización
    If Not ValidarFilaEnBloque(fila) Then Exit Sub

    texto = Trim(CStr(Cells(fila, "A").Value))
    textoNorm = Replace(Replace(UCase(texto), ":", ""), " ", "")

    ' 2) Proteger filas de suma
    If InStr(1, textoNorm, "SUMA") > 0 Then
        MsgBox "No se puede borrar una fila de SUMA.", vbCritical, "Operación no permitida"
        Exit Sub
    End If

    ' 3) Proteger filas de encabezado (título de categoría o columnas de tabla)
    If Cells(fila, "A").MergeArea.Columns.Count > 1 Or _
       InStr(textoNorm, "DESCRIP") > 0 Or InStr(textoNorm, "PRECIO") > 0 Or _
       InStr(textoNorm, "CANTIDAD") > 0 Or InStr(textoNorm, "ITEM") > 0 Then
        MsgBox "No se puede borrar una fila de encabezado o título de categoría.", vbCritical, "Operación no permitida"
        Exit Sub
    End If

    ' 4) Confirmar borrado
    respuesta = MsgBox("¿Seguro que deseas borrar esta fila de concepto?", vbYesNo + vbExclamation, "Confirmar eliminación")

    If respuesta = vbYes Then
        FijarBotones            ' evitar que borrar la fila mueva o deforme los botones
        Rows(fila).Delete
        RestaurarPosicionesBotones
        Call suma               ' recalcular totales inmediatamente
    End If

End Sub
