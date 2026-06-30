' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub AgregarFilaItem()

    Dim filaBase As Long
    Dim filaDestino As Long

    filaBase = ActiveCell.Row
    filaDestino = filaBase + 1

    Rows(filaBase).Copy
    Rows(filaDestino).Insert Shift:=xlDown

End Sub
