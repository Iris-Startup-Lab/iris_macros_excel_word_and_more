' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado con validación de contexto de fila
Sub AgregarFilaFecha()

    Dim filaBase As Long
    Dim filaDestino As Long

    filaBase = ActiveCell.Row
    If Not ValidarFilaEnBloque(filaBase) Then Exit Sub

    FijarBotones            ' los botones no se mueven al insertar la fila

    filaDestino = filaBase + 1

    Rows(filaDestino).Insert Shift:=xlDown

    Sheets("Plantillas").Rows("10:10").Copy

    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormats

    Application.CutCopyMode = False

    RestaurarPosicionesBotones

End Sub
