' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub AgregarFilaItem()

    Dim filaBase As Long
    Dim filaDestino As Long

    ' Importante: Rows(x).Copy + Insert DUPLICA los objetos anclados a esa fila.
    ' FijarBotones los deja libres (xlFreeFloating) y así no se clonan.
    FijarBotones

    filaBase = ActiveCell.Row
    filaDestino = filaBase + 1

    Rows(filaBase).Copy
    Rows(filaDestino).Insert Shift:=xlDown

    Application.CutCopyMode = False

    RestaurarPosicionesBotones

End Sub
