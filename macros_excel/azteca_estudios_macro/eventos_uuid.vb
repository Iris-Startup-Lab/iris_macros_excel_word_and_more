' Autor: Fernando Dorantes Nieto
'
' EVENTOS para automatizar el UUID. OJO: este código NO va en un módulo normal.
' Se pega en dos objetos del proyecto VBA (ventana izquierda del editor):
'
'   BLOQUE A -> objeto "ThisWorkbook"  (doble clic en ThisWorkbook)
'   BLOQUE B -> objeto de la hoja "Cotización " (doble clic en esa hoja)
'
' Con esto, el UUID se asigna al abrir el libro y se revisa cada vez que se
' escribe el nombre del proyecto (C4).

' ============================================================
'  BLOQUE A  ->  pegar dentro del objeto ThisWorkbook
' ============================================================
Private Sub Workbook_Open()
    On Error Resume Next
    AsegurarUUIDCotizacion
End Sub


' ============================================================
'  BLOQUE B  ->  pegar dentro del objeto de la hoja "Cotización "
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Si se modifica C4 (nombre del proyecto), revisamos/reasignamos el UUID.
    If Intersect(Target, Me.Range("C4")) Is Nothing Then Exit Sub

    On Error GoTo salir
    Application.EnableEvents = False
    AsegurarUUIDCotizacion
salir:
    Application.EnableEvents = True
End Sub
