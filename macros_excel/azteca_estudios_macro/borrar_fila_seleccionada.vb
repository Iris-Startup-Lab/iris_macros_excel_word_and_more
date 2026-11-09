' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado para validar contexto de celda activa, proteger encabezados y recalcular totales
' 2026-11-09: usa EsFilaEstructural (M_Blindaje) con coincidencia exacta; antes
'             un concepto que contuviera "SUMA" o "ITEM" no se podía borrar.
Sub BorrarFilaSeleccionada()

    Dim ws As Worksheet
    Dim fila As Long
    Dim respuesta As VbMsgBoxResult

    fila = ActiveCell.Row

    ' 1) Validar que la fila esté dentro del bloque de conceptos de la Cotización
    If Not ValidarFilaEnBloque(fila) Then Exit Sub
    Set ws = HojaCotizacion()

    ' 2) Proteger filas de SUMA TOTAL, títulos de categoría y encabezados
    If EsFilaEstructural(ws, fila) Then
        MsgBox "No se puede borrar una fila de SUMA, de encabezado o título de categoría.", _
               vbCritical, "Operación no permitida"
        Exit Sub
    End If

    ' 3) Confirmar borrado
    respuesta = MsgBox("¿Seguro que deseas borrar esta fila de concepto?", vbYesNo + vbExclamation, "Confirmar eliminación")

    If respuesta = vbYes Then
        FijarBotones ws         ' evitar que borrar la fila mueva o deforme los botones
        ws.Rows(fila).Delete
        RestaurarPosicionesBotones
        Call suma               ' recalcular totales inmediatamente
    End If

End Sub
