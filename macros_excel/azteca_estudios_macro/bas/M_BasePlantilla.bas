Attribute VB_Name = "M_BasePlantilla"
' MÓDULO: M_BasePlantilla
' Propósito: guardar una copia limpia de la hoja Cotización en una hoja muy
' oculta (_BaseAE) y restaurar la cotización desde esa base cuando se requiera
' una nueva cotización.
'
' Depende de:
'   HojaCotizacion(), FijarBotones(), DesbloquearBotones(),
'   BloquearBotones(), RestaurarPosicionesBotones(),
'   GuardarPosicionesBotones(), DesprotegerEstructura(),
'   ProtegerEstructura(), gCreandoHojaSistema (M_Blindaje)
'   suma (Módulo9)
'   ResaltarDetalle() (M_Esquema)

Option Explicit

Private Const NOMBRE_BASE As String = "_BaseAE"

Public Function VersionMBasePlantilla() As String
    VersionMBasePlantilla = "2026-11-09-COMMENTS-LEFT"
End Function

' ============================================================
'  CREACIÓN DE LA BASE (una sola vez)
' ============================================================
Public Sub CrearBasePlantilla()
    Dim ws As Worksheet, base As Worksheet, existente As Worksheet
    Dim ultima As Long, r As Long
    Dim respuesta As VbMsgBoxResult
    Dim pantalla As Boolean, eventos As Boolean, alertas As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "Base de plantilla"
        Exit Sub
    End If

    Set existente = HojaBase()
    If Not existente Is Nothing Then
        respuesta = MsgBox("Ya existe la hoja base " & NOMBRE_BASE & "." & vbCrLf & vbCrLf & _
                           "¿Deseas reemplazarla con la cotización actual?", _
                           vbYesNo + vbExclamation, "Base de plantilla")
        If respuesta <> vbYes Then Exit Sub
    End If

    respuesta = MsgBox("¿Crear la hoja base desde la cotización actual?" & vbCrLf & vbCrLf & _
                       "Después podrás limpiarla manualmente y ocultarla.", _
                       vbYesNo + vbQuestion, "Base de plantilla")
    If respuesta <> vbYes Then Exit Sub

    pantalla = Application.ScreenUpdating
    eventos = Application.EnableEvents
    alertas = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo salir

    gCreandoHojaSistema = True
    DesprotegerEstructura

    If Not existente Is Nothing Then existente.Delete

    ultima = UltimaFila(ws)
    If ultima < 1 Then
        MsgBox "La hoja Cotización está vacía.", vbExclamation, "Base de plantilla"
        GoTo salir
    End If

    Set base = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    base.Name = NOMBRE_BASE

    ws.Range("A1:N" & ultima).Copy
    base.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    ws.Range("A:N").Copy
    base.Range("A:N").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    For r = 1 To ultima
        base.Rows(r).RowHeight = ws.Rows(r).RowHeight
    Next r

    base.Visible = xlSheetVisible
    base.Activate
    Application.Goto base.Range("A1"), True

    gCreandoHojaSistema = False
    ProtegerEstructura

    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos
    Application.DisplayAlerts = alertas

    MsgBox "Hoja base " & NOMBRE_BASE & " creada y visible." & vbCrLf & vbCrLf & _
           "Limpia manualmente los datos que no quieras conservar." & vbCrLf & _
           "Mantén encabezados, categorías, tablas, fórmulas y filas SUMA TOTAL." & vbCrLf & vbCrLf & _
           "Cuando esté lista, ejecuta OcultarBasePlantilla.", _
           vbInformation, "Base de plantilla"
    Exit Sub

salir:
    gCreandoHojaSistema = False
    On Error Resume Next
    ProtegerEstructura
    Application.CutCopyMode = False
    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos
    Application.DisplayAlerts = alertas
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear la base: " & Err.Description, vbCritical, "Base de plantilla"
    End If
End Sub

' ============================================================
'  OCULTAR LA BASE
' ============================================================
Public Sub OcultarBasePlantilla()
    Dim base As Worksheet, ws As Worksheet
    Dim respuesta As VbMsgBoxResult
    Dim pantalla As Boolean, eventos As Boolean

    Set base = HojaBase()
    If base Is Nothing Then
        MsgBox "No existe la hoja base " & NOMBRE_BASE & ".", vbExclamation, "Base de plantilla"
        Exit Sub
    End If

    respuesta = MsgBox("¿La hoja " & NOMBRE_BASE & " ya está limpia y verificada?" & vbCrLf & vbCrLf & _
                       "Se ocultará y no aparecerá en las pestañas.", _
                       vbYesNo + vbQuestion, "Base de plantilla")
    If respuesta <> vbYes Then Exit Sub

    pantalla = Application.ScreenUpdating
    eventos = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo salir

    gCreandoHojaSistema = True
    DesprotegerEstructura
    base.Visible = xlSheetVeryHidden
    ProtegerEstructura
    gCreandoHojaSistema = False

    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos

    Set ws = HojaCotizacion()
    If Not ws Is Nothing Then ws.Activate

    MsgBox "La hoja base " & NOMBRE_BASE & " quedó oculta." & vbCrLf & _
           "La plantilla está lista para crear nuevas cotizaciones.", _
           vbInformation, "Base de plantilla"
    Exit Sub

salir:
    gCreandoHojaSistema = False
    On Error Resume Next
    ProtegerEstructura
    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "No se pudo ocultar la base: " & Err.Description, vbCritical, "Base de plantilla"
    End If
End Sub

' ============================================================
'  RESTAURAR LA COTIZACIÓN DESDE LA BASE
' ============================================================
Public Sub RestaurarBasePlantilla()
    Dim ws As Worksheet, base As Worksheet
    Dim baseLast As Long, wsLast As Long, r As Long
    Dim respuesta As VbMsgBoxResult
    Dim rutaRespaldo As String
    Dim c3 As Variant, c4 As Variant, k3 As Variant, k4 As Variant
    Dim n4 As Variant, h3 As Variant, h4 As Variant
    Dim pantalla As Boolean, eventos As Boolean, alertas As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "Nueva cotización"
        Exit Sub
    End If

    Set base = HojaBase()
    If base Is Nothing Then
        MsgBox "No existe la hoja base " & NOMBRE_BASE & "." & vbCrLf & vbCrLf & _
               "Primero ejecuta CrearBasePlantilla y limpia la base.", _
               vbExclamation, "Nueva cotización"
        Exit Sub
    End If

    respuesta = MsgBox("¿Crear una nueva cotización desde la plantilla base?" & vbCrLf & vbCrLf & _
                       "Se conservan Cliente, Proyecto, Fecha, Versión, ID y UUID." & vbCrLf & _
                       "Los bloques extra se eliminarán.", _
                       vbYesNo + vbQuestion, "Nueva cotización")
    If respuesta <> vbYes Then Exit Sub

    rutaRespaldo = CrearRespaldoRotativo()
    If Len(rutaRespaldo) = 0 Then Exit Sub

    c3 = ws.Range("C3").Value
    c4 = ws.Range("C4").Value
    k3 = ws.Range("K3").Value
    k4 = ws.Range("K4").Value
    n4 = ws.Range("N4").Value
    h3 = ws.Range("H3").Value
    h4 = ws.Range("H4").Value

    pantalla = Application.ScreenUpdating
    eventos = Application.EnableEvents
    alertas = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo salir

    FijarBotones
    DesbloquearBotones ws

    baseLast = UltimaFila(base)
    If baseLast < 1 Then
        MsgBox "La hoja base está vacía.", vbExclamation, "Nueva cotización"
        GoTo salir
    End If

    wsLast = UltimaFila(ws)
    If wsLast > baseLast Then
        ws.Rows((baseLast + 1) & ":" & wsLast).Delete
    End If

    ' Mostrar temporalmente la base para copiar sin problemas.
    gCreandoHojaSistema = True
    DesprotegerEstructura
    base.Visible = xlSheetVisible

    base.Range("A1:N" & baseLast).Copy
    ws.Range("A1").PasteSpecial Paste:=xlPasteAll
    Application.CutCopyMode = False

    base.Range("A:N").Copy
    ws.Range("A:N").PasteSpecial Paste:=xlPasteColumnWidths
    Application.CutCopyMode = False

    base.Visible = xlSheetVeryHidden
    ProtegerEstructura
    gCreandoHojaSistema = False

    For r = 1 To baseLast
        ws.Rows(r).RowHeight = base.Rows(r).RowHeight
    Next r

    On Error Resume Next
    ws.Range("O1:O" & baseLast).ClearContents
    On Error GoTo 0

    AplicarComentariosHeader ws

    ws.Range("C3").Value = c3
    ws.Range("C4").Value = c4
    ws.Range("K3").Value = k3
    ws.Range("N4").Value = n4
    ws.Range("H3").Value = h3
    ws.Range("H4").Value = h4
    ws.Range("K4").NumberFormat = "@"
    ws.Range("K4").Value = k4

    AsegurarAvisoA1 ws

    ws.Rows.ClearOutline
    If baseLast >= 1 Then ws.Rows("1:" & baseLast).EntireRow.Hidden = False

    Call suma
    ResaltarDetalle
    RestaurarPosicionesBotones
    BloquearBotones ws

    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos
    Application.DisplayAlerts = alertas

    Application.Goto ws.Range("A1"), True

    MsgBox "Nueva cotización creada desde la plantilla base." & vbCrLf & vbCrLf & _
           "Respaldo rotativo:" & vbCrLf & rutaRespaldo & vbCrLf & vbCrLf & _
           "Se conservaron Cliente, Proyecto, Fecha, Versión, ID y UUID.", _
           vbInformation, "Nueva cotización"
    Exit Sub

salir:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.ScreenUpdating = pantalla
    Application.EnableEvents = eventos
    Application.DisplayAlerts = alertas
    If Not base Is Nothing Then
        gCreandoHojaSistema = True
        DesprotegerEstructura
        base.Visible = xlSheetVeryHidden
        ProtegerEstructura
        gCreandoHojaSistema = False
    End If
    RestaurarPosicionesBotones
    BloquearBotones ws
    On Error GoTo 0
    If Err.Number <> 0 Then
        MsgBox "No se pudo restaurar la plantilla: " & Err.Description, vbCritical, "Nueva cotización"
    End If
End Sub

' ============================================================
'  BOTÓN
' ============================================================
Public Sub AsignarBotonNuevaCotizacion()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim encontrado As Boolean
    Dim msgError As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "Nueva cotización"
        Exit Sub
    End If

    ws.Activate
    DesbloquearBotones ws
    On Error GoTo salir

    For Each shp In ws.Shapes
        If shp.Type = 8 Then
            If StrComp(shp.Name, "btnLimpiarDatos", vbTextCompare) = 0 Or _
               StrComp(shp.Name, "btnNuevaCotizacion", vbTextCompare) = 0 Or _
               InStr(1, CStr(shp.AlternativeText), "Limpiar datos", vbTextCompare) > 0 Or _
               InStr(1, CStr(shp.AlternativeText), "Nueva cotización", vbTextCompare) > 0 Then
                shp.OnAction = "RestaurarBasePlantilla"
                shp.TextFrame.Characters.Text = "Nueva cotización"
                shp.AlternativeText = "Nueva cotización"
                encontrado = True
            End If
        End If
    Next shp

    If Not encontrado Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, 974.5, 45, 131, 44)
        With shp
            .Name = "btnNuevaCotizacion"
            .Placement = xlFreeFloating
            .Locked = True
            .OnAction = "RestaurarBasePlantilla"
            .TextFrame.Characters.Text = "Nueva cotización"
            .AlternativeText = "Nueva cotización"
        End With
    End If

    FijarBotones ws
    BloquearBotones ws
    GuardarPosicionesBotones

    MsgBox "Botón 'Nueva cotización' asignado a RestaurarBasePlantilla.", _
           vbInformation, "Nueva cotización"
    Exit Sub

salir:
    msgError = Err.Description
    On Error Resume Next
    BloquearBotones ws
    On Error GoTo 0
    If Len(msgError) > 0 Then
        MsgBox "No se pudo asignar el botón: " & msgError, vbCritical, "Nueva cotización"
    End If
End Sub

Public Sub AcomodarBotonesInferiores()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim shNueva As Shape, shReset As Shape, shIzq As Shape, shDer As Shape
    Dim topInferior As Double
    Dim msgError As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "Nueva cotización"
        Exit Sub
    End If

    For Each shp In ws.Shapes
        If shp.Type = 8 Then
            If InStr(1, shp.OnAction, "RestaurarBasePlantilla", vbTextCompare) > 0 Then
                Set shNueva = shp
            ElseIf InStr(1, shp.OnAction, "ResetMacroDefault", vbTextCompare) > 0 Then
                Set shReset = shp
            ElseIf InStr(1, shp.OnAction, "AgregarProducto", vbTextCompare) > 0 Then
                Set shIzq = shp
            ElseIf InStr(1, shp.OnAction, "DuplicarBloque", vbTextCompare) > 0 Then
                Set shDer = shp
            End If
        End If
    Next shp

    If shNueva Is Nothing Or shReset Is Nothing Or shIzq Is Nothing Or shDer Is Nothing Then
        MsgBox "No se encontraron todos los botones para alinear.", vbExclamation, "Nueva cotización"
        Exit Sub
    End If

    ws.Activate
    DesbloquearBotones ws
    On Error GoTo salir

    topInferior = shReset.Top
    shReset.Left = shIzq.Left
    shReset.Top = topInferior

    shNueva.Width = shReset.Width
    shNueva.Height = shReset.Height
    shNueva.Left = shDer.Left
    shNueva.Top = topInferior

    With shNueva.DrawingObject.Font
        .Name = shReset.DrawingObject.Font.Name
        .Size = shReset.DrawingObject.Font.Size
        .Bold = shReset.DrawingObject.Font.Bold
        .Italic = shReset.DrawingObject.Font.Italic
        .Color = vbRed
    End With

    FijarBotones ws
    GuardarPosicionesBotones
    BloquearBotones ws
    Exit Sub

salir:
    msgError = Err.Description
    On Error Resume Next
    BloquearBotones ws
    On Error GoTo 0
    If Len(msgError) > 0 Then
        MsgBox "No se pudieron acomodar los botones: " & msgError, vbCritical, "Nueva cotización"
    End If
End Sub

Public Sub ConfigurarComentariosHeader()
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    DesbloquearBotones ws
    AplicarComentariosHeader ws
    FijarBotones ws
    BloquearBotones ws

    MsgBox "Ayuda del encabezado configurada: comentarios y avisos de lista.", _
           vbInformation, "Encabezado"
End Sub

Private Sub AplicarComentariosHeader(ByVal ws As Worksheet)
    CrearComentarioHeader ws.Range("C3"), _
        "Cliente:" & vbCrLf & "Captura aquí el nombre del cliente."
    CrearComentarioHeader ws.Range("C4"), _
        "Nombre del proyecto:" & vbCrLf & _
        "Captura aquí el nombre del proyecto." & vbCrLf & _
        "Al escribirlo, se generará automáticamente" & vbCrLf & _
        "el ID Cotización en K4."
    CrearComentarioHeader ws.Range("K3"), _
        "Fecha de cotización:" & vbCrLf & _
        "Este campo solo acepta la fecha" & vbCrLf & _
        "en formato dd/mm/aaaa."
    CrearComentarioHeader ws.Range("K4"), _
        "ID Cotización:" & vbCrLf & _
        "Este campo se genera automáticamente" & vbCrLf & _
        "al capturar el Nombre del proyecto en C4." & vbCrLf & _
        "No lo escribas ni lo borres."

    UbicarComentarioIzquierda ws.Range("K3")
    UbicarComentarioIzquierda ws.Range("K4")

    QuitarComentarioHeader ws.Range("H3")
    QuitarComentarioHeader ws.Range("H4")
    ws.Range("H3").Validation.ShowInput = True
    ws.Range("H4").Validation.ShowInput = True
End Sub

Private Sub QuitarComentarioHeader(ByVal celda As Range)
    On Error Resume Next
    If Not celda.Comment Is Nothing Then celda.Comment.Delete
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub UbicarComentarioIzquierda(ByVal celda As Range)
    On Error Resume Next
    If Not celda.Comment Is Nothing Then
        With celda.Comment.Shape
            .Left = celda.Left - .Width - 12
            .Top = celda.Top
            If .Left < 0 Then .Left = 0
        End With
    End If
    On Error GoTo 0
End Sub

Private Sub CrearComentarioHeader(ByVal celda As Range, ByVal texto As String)
    On Error Resume Next

    If Not celda.Comment Is Nothing Then celda.Comment.Delete
    celda.AddComment texto
    If Not celda.Comment Is Nothing Then
        With celda.Comment.Shape
            .Width = 360
            .Height = 180
            .TextFrame.AutoSize = False
            .Visible = False
        End With
    End If
    celda.Validation.ShowInput = False

    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================
'  AUXILIARES
' ============================================================
Private Function HojaBase() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOMBRE_BASE)
    On Error GoTo 0

    Set HojaBase = ws
End Function

Private Function UltimaFila(ByVal ws As Worksheet) As Long
    Dim u As Long

    u = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > u Then
        u = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    End If
    UltimaFila = u
End Function

Private Sub AsegurarAvisoA1(ByVal ws As Worksheet)
    On Error Resume Next

    If Len(Trim$(CStr(ws.Range("A1").Value))) = 0 Then
        ws.Range("A1").Value = "Esta es una plantilla"
    End If

    If ws.Range("A1").Comment Is Nothing Then
        ws.Range("A1").AddComment "Plantilla de cotización" & vbCrLf & _
            "Esta es una PLANTILLA. No borres bloques, filas de categoría ni la fila SUMA TOTAL. " & _
            "Solo modifica las celdas de captura (descripción, fechas, cantidades, precios y descuentos)."
    End If

    Err.Clear
    On Error GoTo 0
End Sub

Private Function CrearRespaldoRotativo() As String
    Dim carpeta As String, nombreBase As String, ruta As String

    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Guarda el libro antes de crear la nueva cotización.", vbExclamation, "Nueva cotización"
        Exit Function
    End If

    carpeta = ThisWorkbook.Path & "\Respaldos"

    On Error Resume Next
    If Dir(carpeta, vbDirectory) = "" Then MkDir carpeta
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear la carpeta de respaldos: " & Err.Description & vbCrLf & _
               "No se restauró nada.", vbCritical, "Nueva cotización"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    nombreBase = ThisWorkbook.Name
    If InStrRev(nombreBase, ".") > 0 Then
        nombreBase = Left$(nombreBase, InStrRev(nombreBase, ".") - 1)
    End If

    ruta = carpeta & "\" & nombreBase & "_ULTIMO_RESPALDO.xlsm"

    If Dir(ruta) <> "" Then Kill ruta
    If Err.Number <> 0 Then
        MsgBox "No se pudo reemplazar el respaldo anterior: " & Err.Description & vbCrLf & _
               "No se restauró nada.", vbCritical, "Nueva cotización"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    ThisWorkbook.SaveCopyAs ruta
    If Err.Number <> 0 Then
        MsgBox "No se pudo crear el respaldo: " & Err.Description & vbCrLf & _
               "No se restauró nada.", vbCritical, "Nueva cotización"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    On Error GoTo 0
    CrearRespaldoRotativo = ruta
End Function
