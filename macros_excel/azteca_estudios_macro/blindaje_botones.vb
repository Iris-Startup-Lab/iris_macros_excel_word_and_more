' Autor: Fernando Dorantes Nieto
'
' MÓDULO: M_Blindaje
' Propósito: blindar la plantilla contra tres fallas operativas reportadas:
'
'   1) LOS BOTONES SE MUEVEN / SE METEN AL ÁREA IMPRIMIBLE.
'      Causa raíz: por omisión un objeto de Excel nace con
'      Placement = xlMoveAndSize, es decir queda ATADO a las celdas que tiene
'      debajo. Todas las macros de estructura hacen Rows(...).Insert /
'      Rows(...).Delete, y además el esquema y el PDF Resumen ocultan filas.
'      Cada una de esas operaciones arrastra el botón, lo aplasta o —cuando la
'      inserción viene de un Rows(x).Copy— lo DUPLICA. De ahí que "se muevan
'      cuando quieren".
'      Solución: anclarlos (xlFreeFloating), guardar su posición correcta en la
'      hoja Meta, poder devolverlos ahí en un clic, y proteger los objetos de
'      dibujo para que nadie los arrastre a mano.
'      OJO: se protegen SOLO los objetos (Contents:=False), así que las celdas
'      siguen editables y ninguna macro existente se rompe.
'
'   2) SE BORRA EL ID COTIZACIÓN (K4)  ->  ver M_UUID.VigilarIDCotizacion.
'
'   3) SE CREAN / COPIAN HOJAS DE COTIZACIÓN  ->  protección de ESTRUCTURA del
'      libro (aquí) + evento Workbook_NewSheet (ver eventos_uuid.vb).
'
' Depende de: HojaCotizacion(), HojaMeta() (módulo M_UUID).

Option Explicit

' Contraseña de blindaje. CÁMBIALA por la de tu equipo antes de repartir el
' archivo (y anótala: sin ella no se puede desproteger la estructura).
Public Const PWD_BLINDAJE As String = "AE2026"

' Bandera que usan las macros para crear hojas de sistema (p. ej. "Meta") sin
' que el evento Workbook_NewSheet las borre.
Public gCreandoHojaSistema As Boolean

' Registro de posiciones dentro de la hoja Meta (columnas libres):
'   D = nombre del objeto | E = Left | F = Top | G = Width | H = Height
Private Const COL_REG As String = "D"
Private Const FILA_REG As Long = 1
Private Const TOLERANCIA As Double = 0.5      ' puntos; por debajo de esto no se toca nada


' ============================================================
'  1) ANCLAJE Y POSICIÓN DE LOS BOTONES
' ============================================================

' Deja todos los objetos de la hoja "libres": ya no se mueven, ni se aplastan,
' ni se duplican cuando se insertan/borran/ocultan filas.
' Es idempotente y no marca el libro como modificado si ya estaba todo anclado.
Public Sub FijarBotones(Optional ByVal ws As Worksheet)
    Dim sh As Shape
    Dim faltaAnclar As Boolean

    If ws Is Nothing Then Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If ws.Shapes.Count = 0 Then Exit Sub

    For Each sh In ws.Shapes
        If sh.Placement <> xlFreeFloating Then faltaAnclar = True: Exit For
    Next sh
    If Not faltaAnclar Then Exit Sub

    DesbloquearBotones ws
    On Error Resume Next
    For Each sh In ws.Shapes
        sh.Placement = xlFreeFloating
        sh.Locked = True
    Next sh
    On Error GoTo 0
    BloquearBotones ws
End Sub

' Toma la foto de "así deben estar los botones". Se corre UNA VEZ, cuando el
' panel de botones esté acomodado como quieres que quede para siempre.
Public Sub GuardarPosicionesBotones()
    Dim ws As Worksheet, meta As Worksheet, sh As Shape
    Dim r As Long, n As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
        Exit Sub
    End If
    Set meta = HojaMeta()

    FijarBotones ws

    meta.Range(COL_REG & FILA_REG & ":H" & FILA_REG + 999).ClearContents

    r = FILA_REG
    For Each sh In ws.Shapes
        meta.Cells(r, COL_REG).Value = sh.Name
        meta.Cells(r, "E").Value = sh.Left
        meta.Cells(r, "F").Value = sh.Top
        meta.Cells(r, "G").Value = sh.Width
        meta.Cells(r, "H").Value = sh.Height
        r = r + 1
        n = n + 1
    Next sh

    MsgBox n & " objeto(s) registrados en su posición actual." & vbCrLf & vbCrLf & _
           "De aquí en adelante «Reacomodar botones» los devolverá exactamente a " & _
           "este lugar, y quedan anclados para que insertar o borrar filas ya no " & _
           "los arrastre.", vbInformation
End Sub

' Devuelve cada objeto registrado a su posición/tamaño de la foto.
' silencioso = True  -> sin mensajes (se llama desde los eventos).
' silencioso = False -> reporta qué reacomodó (botón "Reacomodar botones").
Public Sub RestaurarPosicionesBotones(Optional ByVal silencioso As Boolean = True)
    Dim ws As Worksheet, meta As Worksheet, sh As Shape
    Dim reg As Object
    Dim r As Long, movidos As Long, huerfanos As Long
    Dim L As Double, T As Double, W As Double, H As Double

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    Set meta = HojaMeta()

    Set reg = CreateObject("Scripting.Dictionary")
    r = FILA_REG
    Do While Len(Trim$(CStr(meta.Cells(r, COL_REG).Value))) > 0
        reg(CStr(meta.Cells(r, COL_REG).Value)) = r
        r = r + 1
    Loop

    If reg.Count = 0 Then
        If Not silencioso Then
            MsgBox "Todavía no hay posiciones registradas." & vbCrLf & _
                   "Acomoda los botones a mano una última vez y corre " & _
                   "«GuardarPosicionesBotones».", vbExclamation
        End If
        Exit Sub
    End If

    ' --- Primera pasada: ¿de verdad hay algo que mover? ---
    On Error Resume Next
    For Each sh In ws.Shapes
        If reg.Exists(sh.Name) Then
            r = reg(sh.Name)
            L = CDbl(meta.Cells(r, "E").Value)
            T = CDbl(meta.Cells(r, "F").Value)
            W = CDbl(meta.Cells(r, "G").Value)
            H = CDbl(meta.Cells(r, "H").Value)
            If Abs(sh.Left - L) > TOLERANCIA Or Abs(sh.Top - T) > TOLERANCIA Or _
               Abs(sh.Width - W) > TOLERANCIA Or Abs(sh.Height - H) > TOLERANCIA Or _
               sh.Placement <> xlFreeFloating Then
                movidos = movidos + 1
            End If
        Else
            huerfanos = huerfanos + 1
        End If
    Next sh
    On Error GoTo 0

    If movidos = 0 Then
        BloquearBotones ws
        If Not silencioso Then
            MsgBox "Los botones están en su lugar." & _
                   IIf(huerfanos > 0, vbCrLf & vbCrLf & huerfanos & " objeto(s) sin " & _
                   "registrar (posibles duplicados). Usa «LimpiarBotonesDuplicados» " & _
                   "para revisarlos.", ""), vbInformation
        End If
        Exit Sub
    End If

    ' --- Segunda pasada: reacomodar ---
    Application.ScreenUpdating = False
    DesbloquearBotones ws
    On Error Resume Next
    For Each sh In ws.Shapes
        If reg.Exists(sh.Name) Then
            r = reg(sh.Name)
            sh.Placement = xlFreeFloating
            sh.Left = CDbl(meta.Cells(r, "E").Value)
            sh.Top = CDbl(meta.Cells(r, "F").Value)
            sh.Width = CDbl(meta.Cells(r, "G").Value)
            sh.Height = CDbl(meta.Cells(r, "H").Value)
            sh.Locked = True
        End If
    Next sh
    On Error GoTo 0
    BloquearBotones ws
    Application.ScreenUpdating = True

    If Not silencioso Then
        MsgBox movidos & " objeto(s) devueltos a su posición registrada." & _
               IIf(huerfanos > 0, vbCrLf & vbCrLf & huerfanos & " objeto(s) sin " & _
               "registrar (posibles duplicados). Usa «LimpiarBotonesDuplicados».", ""), _
               vbInformation
    End If
End Sub

' Encuentra los objetos que NO están en el registro: normalmente son los clones
' que generó un Rows(x).Copy + Insert. Pregunta antes de borrar nada.
Public Sub LimpiarBotonesDuplicados()
    Dim ws As Worksheet, meta As Worksheet, sh As Shape
    Dim reg As Object, lista As String
    Dim r As Long, n As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    Set meta = HojaMeta()

    Set reg = CreateObject("Scripting.Dictionary")
    r = FILA_REG
    Do While Len(Trim$(CStr(meta.Cells(r, COL_REG).Value))) > 0
        reg(CStr(meta.Cells(r, COL_REG).Value)) = True
        r = r + 1
    Loop
    If reg.Count = 0 Then
        MsgBox "No hay registro de botones todavía. Corre primero " & _
               "«GuardarPosicionesBotones».", vbExclamation
        Exit Sub
    End If

    For Each sh In ws.Shapes
        If Not reg.Exists(sh.Name) Then
            n = n + 1
            lista = lista & vbCrLf & "  - " & sh.Name
        End If
    Next sh

    If n = 0 Then
        MsgBox "No hay objetos duplicados o de más en la hoja.", vbInformation
        Exit Sub
    End If

    If MsgBox("Se encontraron " & n & " objeto(s) que no están en el registro:" & _
              vbCrLf & lista & vbCrLf & vbCrLf & "¿Eliminarlos?", _
              vbYesNo + vbExclamation, "Objetos de más") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    DesbloquearBotones ws
    On Error Resume Next
    For r = ws.Shapes.Count To 1 Step -1
        If Not reg.Exists(ws.Shapes(r).Name) Then ws.Shapes(r).Delete
    Next r
    On Error GoTo 0
    BloquearBotones ws
    Application.ScreenUpdating = True

    MsgBox n & " objeto(s) eliminados.", vbInformation
End Sub

' Bloquea SOLO los objetos de dibujo: el usuario ya no puede arrastrar, cambiar
' de tamaño ni borrar los botones, pero las celdas siguen 100% editables
' (Contents:=False) y las macros de insertar/borrar filas siguen funcionando.
Public Sub BloquearBotones(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If ws.ProtectDrawingObjects Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=PWD_BLINDAJE, DrawingObjects:=True, _
               Contents:=False, Scenarios:=False
    On Error GoTo 0
End Sub

Public Sub DesbloquearBotones(Optional ByVal ws As Worksheet)
    If ws Is Nothing Then Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectDrawingObjects Then Exit Sub
    On Error Resume Next
    ws.Unprotect Password:=PWD_BLINDAJE
    On Error GoTo 0
End Sub


' ============================================================
'  3) ESTRUCTURA DEL LIBRO (no más hojas de cotización)
' ============================================================

' Con la estructura protegida, Excel deshabilita: Insertar hoja, Eliminar hoja,
' Mover o copiar hoja, Cambiar nombre, Ocultar/Mostrar hoja. Es la forma NATIVA
' y la más difícil de saltarse.
Public Sub ProtegerEstructura()
    On Error Resume Next
    If ThisWorkbook.ProtectStructure Then Exit Sub
    ThisWorkbook.Protect Password:=PWD_BLINDAJE, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

Public Sub DesprotegerEstructura()
    On Error Resume Next
    If Not ThisWorkbook.ProtectStructure Then Exit Sub
    ThisWorkbook.Unprotect Password:=PWD_BLINDAJE
    On Error GoTo 0
End Sub


' ============================================================
'  PUESTA EN MARCHA / MANTENIMIENTO
' ============================================================

' Corre esto UNA VEZ con el panel de botones ya acomodado como debe quedar.
' Deja el libro blindado y guarda la foto de referencia.
Public Sub BlindarLibro()
    Dim ws As Worksheet

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    FijarBotones ws                  ' 1) anclar todo
    GuardarPosicionesBotones         ' 2) foto de referencia (ya avisa por su cuenta)
    AsegurarUUIDCotizacion           ' 3) ID Cotización en su lugar
    ProtegerEstructura               ' 4) sin hojas nuevas
    Application.ScreenUpdating = True

    MsgBox "Libro blindado:" & vbCrLf & _
           "  - Botones anclados, bloqueados y con posición registrada." & vbCrLf & _
           "  - ID Cotización vigilado (se repone si lo borran)." & vbCrLf & _
           "  - Estructura protegida: no se pueden crear ni copiar hojas." & vbCrLf & vbCrLf & _
           "Guarda el archivo como .xlsm.", vbInformation
End Sub

' Para mantenimiento: quita el blindaje y deja mover botones / hojas a mano.
Public Sub DesblindarLibro()
    DesbloquearBotones
    DesprotegerEstructura
    MsgBox "Blindaje retirado (botones y hojas se pueden mover a mano)." & vbCrLf & _
           "Al cerrar y volver a abrir el libro se aplica solo otra vez." & vbCrLf & _
           "Si moviste botones a propósito, corre «GuardarPosicionesBotones».", vbInformation
End Sub
