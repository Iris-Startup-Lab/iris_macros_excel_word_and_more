' Autor: Fernando Dorantes Nieto
'
' EVENTOS de la plantilla. OJO: este código NO va en un módulo normal.
' Se pega en dos objetos del proyecto VBA (ventana izquierda del editor):
'
'   BLOQUE A -> objeto "ThisWorkbook"          (doble clic en ThisWorkbook)
'   BLOQUE B -> objeto de la hoja "Cotización " (doble clic en esa hoja)
'
' Cubre las tres reglas de blindaje:
'   1) Los botones quedan anclados y vuelven a su sitio solos.
'   2) El ID Cotización (K4) se repone si lo borran o lo editan a mano.
'   3) No se pueden crear ni copiar hojas de cotización.
'
' Requiere los módulos: M_UUID (asignar_uuid_cotizacion) y
'                       M_Blindaje (blindaje_botones).

' ============================================================
'  BLOQUE A  ->  pegar dentro del objeto ThisWorkbook
'                (REEMPLAZA por completo lo que hoy tiene)
' ============================================================
Private Sub Workbook_Open()
    On Error Resume Next
    AsegurarUUIDCotizacion            ' (2) el ID siempre presente al abrir
    FijarBotones                      ' (1) anclar objetos: no se mueven con las celdas
    RestaurarPosicionesBotones        ' (1) si alguno viajó, regresa a su posición
    ProtegerEstructura                ' (3) sin insertar / copiar / mover hojas
End Sub

' (3) Última barrera: si de alguna forma se logra crear una hoja (p. ej. pegando
' desde otro libro, o si alguien desprotegió la estructura), se elimina al vuelo.
' NOTA: este bloque se copia y pega a mano, así que el mensaje se arma línea por
' línea (sin continuaciones "_") para que un copiado que pierda comillas falle de
' forma obvia y aislada, en lugar de romper todo el Sub.
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    Dim nombre As String
    Dim msg As String

    ' Las macros crean hojas de SISTEMA ("Meta") con esta bandera en True.
    If gCreandoHojaSistema Then Exit Sub

    nombre = Sh.Name

    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error Resume Next
    Sh.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ProtegerEstructura

    msg = "No se permite crear hojas nuevas ni copiar la hoja de Cotización."
    msg = msg & vbCrLf & "Se eliminó la hoja: " & nombre
    msg = msg & vbCrLf & vbCrLf & "Para una cotización nueva, cierra este archivo"
    msg = msg & " y trabaja sobre una COPIA del .xlsm."
    msg = msg & vbCrLf & "Un archivo = una cotización = un ID."
    MsgBox msg, vbCritical, "Hoja no permitida"
End Sub

' Red de seguridad al guardar: el archivo guardado siempre lleva su ID y sus
' botones en su lugar.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    RestaurarIDSiFalta
    RestaurarPosicionesBotones
    ProtegerEstructura
End Sub

' Los botones no deben salir en la impresión ni en los PDFs: si alguno se colara
' al área imprimible, esto lo devuelve a su panel antes de imprimir.
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    On Error Resume Next
    RestaurarIDSiFalta
    RestaurarPosicionesBotones
End Sub


' ============================================================
'  BLOQUE B  ->  pegar dentro del objeto de la hoja "Cotización "
'                (REEMPLAZA por completo lo que hoy tiene)
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' (2) Vigila C4 (proyecto) y K4 (ID Cotización). Si borran o alteran K4, se
    ' reimprime el MISMO UUID guardado en la hoja Meta.
    On Error Resume Next
    VigilarIDCotizacion Target
End Sub

Private Sub Worksheet_Activate()
    ' (1) Cada vez que se entra a la hoja se verifica el anclaje y la posición.
    ' No hace nada (ni marca el libro como modificado) si todo está en su sitio.
    On Error Resume Next
    FijarBotones Me
    RestaurarPosicionesBotones
End Sub
