' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
' Modificado con validación de contexto de fila y recálculo automático
Sub AgregarFilaItem()

    Dim filaBase As Long
    Dim filaDestino As Long

    filaBase = ActiveCell.Row
    If Not ValidarFilaEnBloque(filaBase) Then Exit Sub

    ' Importante: Rows(x).Copy + Insert DUPLICA los objetos anclados a esa fila.
    ' FijarBotones los deja libres (xlFreeFloating) y así no se clonan.
    FijarBotones

    filaDestino = filaBase + 1

    Rows(filaBase).Copy
    Rows(filaDestino).Insert Shift:=xlDown

    Application.CutCopyMode = False

    RestaurarPosicionesBotones
    Call suma               ' recalcular totales

End Sub
