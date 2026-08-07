' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub AgregarFilaFecha()

    Dim filaDestino As Long

    FijarBotones            ' los botones no se mueven al insertar la fila

    filaDestino = ActiveCell.Row + 1

    Rows(filaDestino).Insert Shift:=xlDown

    Sheets("Plantillas").Rows("10:10").Copy

    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormats

    Application.CutCopyMode = False

    RestaurarPosicionesBotones

End Sub


