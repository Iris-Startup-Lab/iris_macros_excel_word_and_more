' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado con validación de contexto de fila y recálculo automático
' 2026-11-09: rechaza filas estructurales (SUMA, títulos, encabezados) para no
'             duplicar una SUMA TOTAL y descuadrar los totales.
Sub AgregarFilaItem()

    Dim ws As Worksheet
    Dim filaBase As Long
    Dim filaDestino As Long

    filaBase = ActiveCell.Row
    If Not ValidarFilaEnBloque(filaBase) Then Exit Sub

    Set ws = HojaCotizacion()
    If EsFilaEstructural(ws, filaBase) Then
        MsgBox "Selecciona una fila de concepto (no una fila de SUMA, título de " & _
               "categoría o encabezado de columnas) para agregar un ítem debajo.", _
               vbExclamation, "Fila no válida"
        Exit Sub
    End If

    ' Importante: Rows(x).Copy + Insert DUPLICA los objetos anclados a esa fila.
    ' FijarBotones los deja libres (xlFreeFloating) y así no se clonan.
    FijarBotones ws

    filaDestino = filaBase + 1

    ws.Rows(filaBase).Copy
    ws.Rows(filaDestino).Insert Shift:=xlDown

    Application.CutCopyMode = False

    RestaurarPosicionesBotones
    Call suma               ' recalcular totales

End Sub
