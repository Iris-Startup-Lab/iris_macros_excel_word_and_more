' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado con validación de contexto de fila y recálculo automático
Sub AgregarFilaProducto()

    Dim filaBase As Long
    Dim filaDestino As Long

    filaBase = ActiveCell.Row
    If Not ValidarFilaEnBloque(filaBase) Then Exit Sub

    FijarBotones            ' los botones no se mueven al insertar la fila

    filaDestino = filaBase + 1

    Rows(filaDestino).Insert Shift:=xlDown

    Sheets("Plantillas").Rows("11:11").Copy

    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormats
    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormulas

    Application.CutCopyMode = False

    RestaurarPosicionesBotones
    Call suma               ' recalcular totales

End Sub
