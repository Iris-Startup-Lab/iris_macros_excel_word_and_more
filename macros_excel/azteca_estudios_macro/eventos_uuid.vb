' Autores: Fernando Dorantes Nieto
' Diana Berumen Estrada
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
' 2026-11-09:
'   - Detección de archivo copiado (VerificarCopiaDeArchivo) al abrir y
'     después de "Guardar como" (Workbook_AfterSave).
'   - Los errores ya no se silencian del todo: se anotan en Meta!J:M con
'     RegistrarError y el evento continúa con la siguiente instrucción.
'
' Requiere los módulos: M_UUID (asignar_uuid_cotizacion),
'                       M_Blindaje (blindaje_botones) y
'                       M_Catalogos (catalogo_categorias).

' ============================================================
'  BLOQUE A  ->  pegar dentro del objeto ThisWorkbook
'                (REEMPLAZA por completo lo que hoy tiene)
' ============================================================
Private Sub Workbook_Open()
    On Error GoTo fallo
    VerificarCopiaDeArchivo           ' (2) ¿es una copia de otra cotización?
    AsegurarUUIDCotizacion            ' (2) el ID siempre presente al abrir
    FijarBotones                      ' (1) anclar objetos: no se mueven con las celdas
    RestaurarPosicionesBotones        ' (1) si alguno viajó, regresa a su posición
    ProtegerEstructura                ' (3) sin insertar / copiar / mover hojas
    Exit Sub
fallo:
    RegistrarError "Workbook_Open", Err.Number, Err.Description
    Resume Next
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
    On Error GoTo fallo
    RestaurarIDSiFalta
    RestaurarPosicionesBotones
    ProtegerEstructura
    Exit Sub
fallo:
    RegistrarError "Workbook_BeforeSave", Err.Number, Err.Description
    Resume Next
End Sub

' Tras un "Guardar como" el nombre del archivo cambia: se pregunta en ese
' momento si es una cotización nueva y, si hubo cambios, se vuelve a guardar.
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    On Error GoTo fallo
    If Success Then VerificarCopiaDeArchivo True
    Exit Sub
fallo:
    RegistrarError "Workbook_AfterSave", Err.Number, Err.Description
End Sub

' Los botones no deben salir en la impresión ni en los PDFs: si alguno se colara
' al área imprimible, esto lo devuelve a su panel antes de imprimir.
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    On Error GoTo fallo
    RestaurarIDSiFalta
    RestaurarPosicionesBotones
    Exit Sub
fallo:
    RegistrarError "Workbook_BeforePrint", Err.Number, Err.Description
    Resume Next
End Sub


' ============================================================
'  BLOQUE B  ->  pegar dentro del objeto de la hoja "Cotización "
'                (REEMPLAZA por completo lo que hoy tiene)
' ============================================================
Private Sub Worksheet_Change(ByVal Target As Range)
    ' (2) Vigila C4 (proyecto) y K4 (ID Cotización). Si borran o alteran K4, se
    ' reimprime el MISMO UUID guardado en la hoja Meta.
    ' (4) Si se elige un foro por la flecha nativa, agrega también su descripción.
    ' (6) Recalcula subtotales y totales al editar o vaciar celdas de productos.
    On Error GoTo fallo
    VigilarIDCotizacion Target
    AjustarDescripcionForo Target
    RecalcularTrasCambio Target
    Exit Sub
fallo:
    RegistrarError "Worksheet_Change", Err.Number, Err.Description
    Application.EnableEvents = True
    Resume Next
End Sub

Private Sub Worksheet_Activate()
    ' (1) Cada vez que se entra a la hoja se verifica el anclaje y la posición.
    ' No hace nada (ni marca el libro como modificado) si todo está en su sitio.
    ' (3) También limpia cualquier combo temporal huérfano del catálogo.
    On Error GoTo fallo
    LimpiarCatalogoTemporal
    FijarBotones Me
    RestaurarPosicionesBotones
    Exit Sub
fallo:
    RegistrarError "Worksheet_Activate", Err.Number, Err.Description
    Resume Next
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    ' (4) Catálogo de Foros: doble clic en una fila de foro abre el combo de Espacios.
    ' Las categorías usan directamente la validación nativa de la celda A:N.
    On Error GoTo fallo
    If EsFilaForoEn(Target) Then
        Cancel = True
        AbrirCatalogoForoEn Target
    End If
    Exit Sub
fallo:
    RegistrarError "Worksheet_BeforeDoubleClick", Err.Number, Err.Description
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    ' (3) Si el usuario se mueve fuera del encabezado, se cierra el catálogo temporal.
    On Error GoTo fallo
    LimpiarCatalogoSiCorresponde Target
    Exit Sub
fallo:
    RegistrarError "Worksheet_SelectionChange", Err.Number, Err.Description
End Sub

Private Sub Worksheet_Deactivate()
    ' (3) Si cambia de hoja, se limpia cualquier catálogo temporal abierto.
    On Error GoTo fallo
    LimpiarCatalogoTemporal
    Exit Sub
fallo:
    RegistrarError "Worksheet_Deactivate", Err.Number, Err.Description
End Sub
