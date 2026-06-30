' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub AgregarFilaProducto()

    Dim filaDestino As Long

    filaDestino = ActiveCell.Row + 1

    Rows(filaDestino).Insert Shift:=xlDown

    Sheets("Plantillas").Rows("11:11").Copy

    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormats
    Rows(filaDestino).PasteSpecial Paste:=xlPasteFormulas

    Application.CutCopyMode = False

End Sub
