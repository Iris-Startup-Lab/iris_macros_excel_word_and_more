' Autor: Fernando Dorantes Nieto
'
' MÓDULO: M_UUID
' Propósito: asignar automáticamente un UUID (ID Cotización, celda K4) que:
'   - Es único por cotización.
'   - Queda ATADO al nombre del proyecto (celda C4). Solo se regenera si C4 cambia.
'   - Se MANTIENE aunque cambie la versión (N4) o se reabra el archivo.
'   - Es la misma cadena que luego se usa en el nombre del archivo PDF, de modo
'     que archivo e "ID Cotización" siempre coinciden.
'
' El par (proyecto, uuid) se guarda en una hoja MUY oculta llamada "Meta",
' que la propia macro crea la primera vez. Así viaja dentro del libro.
'
' Depende de la función GENERAR_UUID() (módulo macro_generadora_uuid.vb).

Option Explicit

' --- Configuración compartida por los demás módulos ---
Public Const COL_MARCA As String = "O"       ' columna oculta con la marca de detalle
Public Const MARCA_DETALLE As String = "D"    ' valor que marca una fila como "detalle colapsable"
Public Const FILA_INICIO As Long = 7          ' primera fila de bloques de categorías

' --- Helpers de localización de hojas (robustos ante espacios/acentos) ---
' 2026-11-09: ahora sí es insensible a acentos y espacios (antes solo a
' mayúsculas), y da prioridad a la coincidencia EXACTA antes que a la parcial,
' para que "Cotización" gane sobre p. ej. "Plantilla Cotización".
' Las hojas de sistema (Meta, _BaseAE...) nunca se devuelven.
Public Function HojaQueContiene(ByVal token As String) As Worksheet
    Dim ws As Worksheet
    Dim t As String, nombre As String

    t = NormalizarNombreHoja(token)
    If Len(t) = 0 Then Exit Function

    ' 1) Coincidencia exacta
    For Each ws In ThisWorkbook.Worksheets
        If Not EsHojaSistema(ws) Then
            If NormalizarNombreHoja(ws.Name) = t Then
                Set HojaQueContiene = ws
                Exit Function
            End If
        End If
    Next ws

    ' 2) Coincidencia parcial (tolera "Cotización 2", "COTIZACION AE", etc.)
    For Each ws In ThisWorkbook.Worksheets
        If Not EsHojaSistema(ws) Then
            nombre = NormalizarNombreHoja(ws.Name)
            If InStr(1, nombre, t) > 0 Then
                Set HojaQueContiene = ws
                Exit Function
            End If
        End If
    Next ws
End Function

Public Function HojaCotizacion() As Worksheet
    Set HojaCotizacion = HojaQueContiene("COTIZACION")
End Function

Private Function NormalizarNombreHoja(ByVal s As String) As String
    s = UCase$(s)
    s = Replace(s, "Á", "A"): s = Replace(s, "É", "E"): s = Replace(s, "Í", "I")
    s = Replace(s, "Ó", "O"): s = Replace(s, "Ú", "U"): s = Replace(s, "Ü", "U")
    s = Replace(s, " ", ""): s = Replace(s, Chr$(160), "")
    NormalizarNombreHoja = s
End Function

Private Function EsHojaSistema(ByVal ws As Worksheet) As Boolean
    EsHojaSistema = (StrComp(ws.Name, "Meta", vbTextCompare) = 0) Or (Left$(ws.Name, 1) = "_")
End Function

' Pública porque M_Blindaje también guarda ahí el registro de posiciones de los
' botones (columnas D:H).
Public Function HojaMeta() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Meta")
    On Error GoTo 0

    If ws Is Nothing Then
        ' La estructura del libro está protegida (para que nadie cree hojas de
        ' cotización). Aquí sí necesitamos crear una hoja de SISTEMA: avisamos
        ' con la bandera para que Workbook_NewSheet no la borre.
        ' 2026-11-09: con manejador de errores. Antes, si Add fallaba (p. ej.
        ' contraseña distinta), la bandera quedaba en True para siempre y
        ' Workbook_NewSheet dejaba de bloquear hojas nuevas.
        Dim activaAntes As Object
        Set activaAntes = ActiveSheet
        On Error GoTo fallo
        gCreandoHojaSistema = True
        DesprotegerEstructura

        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Meta"
        ws.Range("A1").Value = "PROYECTO"
        ws.Range("A2").Value = "UUID"
        ws.Range("A3").Value = "ACTUALIZADO"
        ws.Range("A4").Value = "ARCHIVO"
        ws.Range("D1").Value = ""        ' D:H = registro de posiciones de botones
        ws.Visible = xlSheetVeryHidden
        If Not activaAntes Is Nothing Then activaAntes.Activate   ' Add activa la hoja nueva
        On Error GoTo 0
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden   ' no aparece ni en el menú de "Mostrar hoja"
    On Error GoTo 0

    If gCreandoHojaSistema Then
        ProtegerEstructura
        gCreandoHojaSistema = False
    End If

    Set HojaMeta = ws
    Exit Function

fallo:
    Dim numErr As Long, descErr As String
    numErr = Err.Number: descErr = Err.Description
    On Error Resume Next
    ' Si la hoja alcanzó a crearse pero no se pudo nombrar, se descarta.
    If Not ws Is Nothing Then
        If ws.Name <> "Meta" Then
            Application.DisplayAlerts = False
            ws.Delete
            Application.DisplayAlerts = True
            Set ws = Nothing
        End If
    End If
    ProtegerEstructura
    gCreandoHojaSistema = False
    On Error GoTo 0
    MsgBox "No se pudo crear la hoja de sistema «Meta» (" & numErr & "):" & vbCrLf & descErr & _
           vbCrLf & vbCrLf & "Verifica que la protección del libro use la contraseña de las macros.", _
           vbCritical, "Hoja Meta"
    Set HojaMeta = ws
End Function

' UUID que el libro tiene guardado para el proyecto actual, sin generar nada.
' Devuelve "" si todavía no hay ninguno.
Public Function UUIDGuardado() As String
    Dim meta As Worksheet
    Set meta = HojaMeta()
    If Not meta Is Nothing Then UUIDGuardado = Trim$(CStr(meta.Range("B2").Value))
End Function

' --- Núcleo: garantiza que K4 tenga el UUID correcto para el proyecto actual ---
Public Sub AsegurarUUIDCotizacion()
    Dim ws As Worksheet, meta As Worksheet
    Dim proyecto As String, proyGuardado As String, uuidGuardado As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    proyecto = Trim$(CStr(ws.Range("C4").Value))
    ' Sin nombre de proyecto todavía no generamos nada (evita UUIDs "huérfanos").
    ' Tampoco se permite un ID escrito a mano: sin proyecto no hay identidad,
    ' así que K4 debe quedar vacío (antes se conservaba el texto inventado y
    ' terminaba en el nombre de los PDFs).
    If Len(proyecto) = 0 Then
        If Len(CStr(ws.Range("K4").Value)) > 0 Then ws.Range("K4").ClearContents
        Exit Sub
    End If

    Set meta = HojaMeta()
    If meta Is Nothing Then Exit Sub
    proyGuardado = Trim$(CStr(meta.Range("B1").Value))
    uuidGuardado = Trim$(CStr(meta.Range("B2").Value))

    ' Regeneramos si: no hay UUID, o el proyecto cambió.
    If Len(uuidGuardado) = 0 Then
        uuidGuardado = GENERAR_UUID()
        meta.Range("B1").Value = proyecto
        meta.Range("B2").Value = uuidGuardado
        meta.Range("B3").Value = Now
        RegistrarArchivoEnMeta meta
    ElseIf StrComp(proyGuardado, proyecto, vbTextCompare) <> 0 Then
        ' El nombre de proyecto cambió. Para evitar perder el folio por una simple
        ' corrección ortográfica o tipográfica, preguntamos al usuario si desea conservar el UUID:
        Dim respUUID As VbMsgBoxResult
        respUUID = MsgBox("El nombre del proyecto ha cambiado:" & vbCrLf & vbCrLf & _
                          "  • Nombre anterior: """ & proyGuardado & """" & vbCrLf & _
                          "  • Nombre nuevo:    """ & proyecto & """" & vbCrLf & vbCrLf & _
                          "¿Deseas generar un nuevo ID de cotización (UUID)?" & vbCrLf & vbCrLf & _
                          "  [Sí]  -> Generar un nuevo folio (cotización nueva)." & vbCrLf & _
                          "  [No]  -> Conservar el ID actual (corrección ortográfica / misma cotización).", _
                          vbYesNo + vbQuestion, "Cambio de Nombre de Proyecto")
        If respUUID = vbYes Then
            uuidGuardado = GENERAR_UUID()
            meta.Range("B1").Value = proyecto
            meta.Range("B2").Value = uuidGuardado
            meta.Range("B3").Value = Now
            RegistrarArchivoEnMeta meta
        Else
            ' Conservar el UUID existente y actualizar solo el nombre en Meta
            meta.Range("B1").Value = proyecto
            meta.Range("B3").Value = Now
        End If
    End If

    ' Meta es la fuente de la verdad: si K4 está vacío (lo borraron) o alguien lo
    ' escribió a mano, se REIMPRIME el mismo UUID. No se inventa uno nuevo.
    ws.Range("K4").NumberFormat = "@"   ' texto, para que Excel no lo deforme
    If CStr(ws.Range("K4").Value) <> uuidGuardado Then
        ws.Range("K4").Value = uuidGuardado
    End If
End Sub

' ============================================================
'  DETECCIÓN DE COPIA DE ARCHIVO (bug 2, 2026-11-09)
' ============================================================
' Al copiar el .xlsm la copia hereda el UUID de Meta: dos cotizaciones con el
' mismo ID. Meta!B4 guarda el NOMBRE del archivo dueño del UUID. Si al abrir el
' nombre no coincide, se pregunta si es una cotización nueva.
' Se compara solo el nombre (no la ruta completa) para no preguntar cada vez que
' el archivo llega por correo, se abre desde OneDrive/SharePoint o se mueve de
' carpeta. Limitación: una copia con el MISMO nombre en otra carpeta no se
' detecta.

Public Sub RegistrarArchivoEnMeta(Optional ByVal meta As Worksheet)
    If meta Is Nothing Then Set meta = HojaMeta()
    If meta Is Nothing Then Exit Sub
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub     ' libro sin guardar: aún no tiene nombre real
    meta.Range("A4").Value = "ARCHIVO"
    If CStr(meta.Range("B4").Value) <> ThisWorkbook.Name Then meta.Range("B4").Value = ThisWorkbook.Name
End Sub

' Se llama desde Workbook_Open ANTES de AsegurarUUIDCotizacion, y desde
' Workbook_AfterSave (con guardarAlTerminar = True) para resolver un
' "Guardar como" en ese momento y volver a guardar el archivo ya corregido.
Public Sub VerificarCopiaDeArchivo(Optional ByVal guardarAlTerminar As Boolean = False)
    Dim eventos As Boolean
    If Not VerificarCopiaInterno() Then Exit Sub
    If Not guardarAlTerminar Then Exit Sub
    eventos = Application.EnableEvents
    On Error Resume Next
    Application.EnableEvents = False     ' sin BeforeSave/AfterSave recursivos
    ThisWorkbook.Save
    Application.EnableEvents = eventos
    On Error GoTo 0
End Sub

' Devuelve True si escribió algo en Meta/K4.
Private Function VerificarCopiaInterno() As Boolean
    Dim ws As Worksheet, meta As Worksheet
    Dim archivoReg As String, uuidReg As String, nuevo As String
    Dim msg As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Function
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If Len(Trim$(CStr(ws.Range("C4").Value))) = 0 Then Exit Function  ' sin proyecto no hay identidad

    Set meta = HojaMeta()
    If meta Is Nothing Then Exit Function
    uuidReg = Trim$(CStr(meta.Range("B2").Value))
    If Len(uuidReg) = 0 Then Exit Function             ' todavía no hay UUID que duplicar
    archivoReg = Trim$(CStr(meta.Range("B4").Value))

    ' Libros creados antes de este cambio: se adopta el nombre actual sin preguntar.
    If Len(archivoReg) = 0 Then
        RegistrarArchivoEnMeta meta
        VerificarCopiaInterno = True
        Exit Function
    End If
    If StrComp(archivoReg, ThisWorkbook.Name, vbTextCompare) = 0 Then Exit Function
    VerificarCopiaInterno = True

    msg = "Este archivo tiene el ID Cotización de otro archivo:" & vbCrLf & vbCrLf & _
          "  • Archivo original: " & archivoReg & vbCrLf & _
          "  • Este archivo:     " & ThisWorkbook.Name & vbCrLf & _
          "  • ID Cotización:    " & uuidReg & vbCrLf & vbCrLf & _
          "¿Es una COTIZACIÓN NUEVA (copia del archivo)?" & vbCrLf & vbCrLf & _
          "  [Sí]  -> Asignar un ID nuevo a este archivo." & vbCrLf & _
          "  [No]  -> Es la misma cotización (solo se renombró o se guardó como)."

    If MsgBox(msg, vbYesNo + vbQuestion, "¿Cotización copiada?") = vbYes Then
        nuevo = GENERAR_UUID()
        meta.Range("B1").Value = Trim$(CStr(ws.Range("C4").Value))
        meta.Range("B2").Value = nuevo
        meta.Range("B3").Value = Now
        RegistrarArchivoEnMeta meta
        Dim eventos As Boolean
        eventos = Application.EnableEvents
        Application.EnableEvents = False
        ws.Range("K4").NumberFormat = "@"
        ws.Range("K4").Value = nuevo
        Application.EnableEvents = eventos
        MsgBox "Nuevo ID Cotización asignado:" & vbCrLf & nuevo, vbInformation, "ID Cotización"
    Else
        RegistrarArchivoEnMeta meta
    End If
End Function

' Vigilante del ID Cotización. Se llama desde Worksheet_Change de la hoja
' Cotización y cubre los dos casos:
'   - Cambió C4 (nombre del proyecto)  -> se asigna el UUID que toca.
'   - Se borró o se editó K4 a mano    -> se REIMPRIME el mismo UUID y se avisa.
' Es reentrante-seguro (Static ocupado) y siempre deja EnableEvents en True.
Public Sub VigilarIDCotizacion(ByVal Target As Range)
    Static ocupado As Boolean
    Dim ws As Worksheet
    Dim antes As String, despues As String
    Dim tocaronID As Boolean

    If ocupado Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub
    If Intersect(Target, ws.Range("C4,K4")) Is Nothing Then Exit Sub

    tocaronID = Not (Intersect(Target, ws.Range("K4")) Is Nothing)

    ocupado = True
    On Error GoTo salir
    Application.EnableEvents = False

    antes = Trim$(CStr(ws.Range("K4").Value))
    AsegurarUUIDCotizacion
    despues = Trim$(CStr(ws.Range("K4").Value))

    Application.EnableEvents = True
    ocupado = False

    ' Solo avisamos cuando el usuario tocó K4 y tuvimos que reponerlo.
    If tocaronID And StrComp(antes, despues, vbTextCompare) <> 0 Then
        If Len(despues) > 0 Then
            MsgBox "El ID Cotización no se edita ni se borra a mano: " & _
                   "se repuso automáticamente." & vbCrLf & vbCrLf & _
                   "ID Cotización: " & despues, vbExclamation, "ID Cotización"
        Else
            MsgBox "No hay ID Cotización porque falta el nombre del proyecto (C4)." & _
                   vbCrLf & "Escribe el proyecto y el ID se generará solo.", _
                   vbExclamation, "ID Cotización"
        End If
    End If
    Exit Sub

salir:
    Application.EnableEvents = True
    ocupado = False
End Sub

' Botón/red de seguridad: repone el ID Cotización si falta, sin preguntar nada.
' Se llama al abrir el libro, al guardar y antes de generar los PDFs.
Public Sub RestaurarIDSiFalta()
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Len(Trim$(CStr(ws.Range("K4").Value))) > 0 Then
        If StrComp(Trim$(CStr(ws.Range("K4").Value)), UUIDGuardado(), vbTextCompare) = 0 Then Exit Sub
    End If
    On Error Resume Next
    Application.EnableEvents = False
    AsegurarUUIDCotizacion
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' Devuelve el UUID actual (asegurándolo primero). Útil para otros módulos.
Public Function UUIDActual() As String
    AsegurarUUIDCotizacion
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If Not ws Is Nothing Then UUIDActual = Trim$(CStr(ws.Range("K4").Value))
End Function

' Botón manual opcional: fuerza revisión/asignación del UUID.
Public Sub AsignarUUID_Click()
    AsegurarUUIDCotizacion
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
    Else
        MsgBox "ID Cotización (UUID): " & vbCrLf & ws.Range("K4").Value, vbInformation
    End If
End Sub
