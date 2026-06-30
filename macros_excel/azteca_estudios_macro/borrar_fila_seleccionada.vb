' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub BorrarFilaSeleccionada()

    Dim fila As Long
    Dim texto As String
    Dim respuesta As VbMsgBoxResult

    fila = ActiveCell.Row
    
    texto = Cells(fila, "A").Value

    If InStr(1, texto, "SUMA TOTAL:") > 0 Then
        MsgBox "No se puede borrar una fila de SUMA.", vbCritical
        Exit Sub
    End If

    respuesta = MsgBox("¿Seguro que deseas borrar esta fila?", vbYesNo + vbExclamation)

    If respuesta = vbYes Then
        Rows(fila).Delete
    End If

End Sub
