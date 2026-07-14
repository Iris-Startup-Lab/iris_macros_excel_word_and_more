Attribute VB_Name = "M_Esquema"
' Autor: Fernando Dorantes Nieto
'
' MÓDULO: M_Esquema
' Propósito: construir un esquema (outline) de 2 niveles sobre la hoja Cotización
' y ofrecer colapsar/expandir a demanda, más el marcado de filas de detalle.
'
'   Nivel 1  -> cuerpo completo de cada categoría (se colapsa la categoría entera,
'               quedando solo su barra de título y su SUMA TOTAL).
'   Nivel 2  -> rachas de filas marcadas con "D" en la columna oculta O
'               (el detalle largo de cada concepto).
'
' Las categorías se detectan comparando la columna A contra la lista de la hoja
' "Categorías". La marca de detalle la pone el usuario con MarcarComoDetalle,
' y las filas marcadas se resaltan en rojizo (formato condicional sobre O="D").
'
' Depende de: HojaCotizacion(), COL_MARCA, MARCA_DETALLE, FILA_INICIO (módulo M_UUID).

Option Explicit

' ============================================================
'  CONSTRUCCIÓN DEL ESQUEMA
' ============================================================
Public Sub ConstruirEsquema()
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False

    On Error Resume Next
    ws.Rows.ClearOutline          ' limpiamos cualquier agrupación previa
    On Error GoTo 0
    ws.Outline.SummaryRow = xlAbove

    Dim cats As Object
    Set cats = CargarCategorias()

    Dim ultima As Long, i As Long
    Dim a As String, inicioCat As Long, hdrEnd As Long
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    inicioCat = 0

    For i = FILA_INICIO To ultima
        a = UCase$(Trim$(CStr(ws.Cells(i, "A").Value)))

        If cats.Exists(a) Then
            inicioCat = i                       ' encabezado de categoría
        ElseIf InStr(1, a, "SUMA TOTAL") > 0 And inicioCat > 0 Then
            hdrEnd = FilaFinEncabezado(ws, inicioCat)
            ' Nivel 1: agrupamos el cuerpo (después del título hasta la fila
            ' anterior a SUMA TOTAL), dejando visible el título y el total.
            If (i - 1) > hdrEnd Then
                ws.Rows(hdrEnd + 1 & ":" & i - 1).Group
                ' Nivel 2: rachas de filas de detalle dentro del cuerpo.
                AgruparDetalle ws, hdrEnd + 1, i - 1
            End If
            inicioCat = 0
        End If
    Next i

    ws.Columns(COL_MARCA).Hidden = True
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=2          ' categorías abiertas, detalle colapsado
    On Error GoTo 0

    ResaltarDetalle                             ' asegura el rojizo de las filas marcadas

    Application.ScreenUpdating = True
    MsgBox "Esquema reconstruido." & vbCrLf & _
           "Usa los botones [-]/[+] del margen izquierdo para colapsar categorías o detalle a demanda.", _
           vbInformation
End Sub

Private Function FilaFinEncabezado(ByVal ws As Worksheet, ByVal r As Long) As Long
    ' El título de categoría suele estar fusionado en 2 filas (ej. A7:N8).
    Dim m As Range
    Set m = ws.Cells(r, "A").MergeArea
    FilaFinEncabezado = m.Row + m.Rows.Count - 1
End Function

Private Sub AgruparDetalle(ByVal ws As Worksheet, ByVal topRow As Long, ByVal bottomRow As Long)
    Dim j As Long, runStart As Long
    runStart = 0
    For j = topRow To bottomRow
        If EsDetalle(ws, j) Then
            If runStart = 0 Then runStart = j
        Else
            If runStart > 0 Then
                ws.Rows(runStart & ":" & j - 1).Group
                runStart = 0
            End If
        End If
    Next j
    If runStart > 0 Then ws.Rows(runStart & ":" & bottomRow).Group
End Sub

Private Function EsDetalle(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    EsDetalle = (UCase$(Trim$(CStr(ws.Cells(r, COL_MARCA).Value))) = UCase$(MARCA_DETALLE))
End Function

Private Function CargarCategorias() As Object
    Dim d As Object, ws As Worksheet, i As Long, ultima As Long, v As String
    Set d = CreateObject("Scripting.Dictionary")
    Set ws = HojaQueContiene("CATEGOR")
    If ws Is Nothing Then
        Set CargarCategorias = d
        Exit Function
    End If
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ultima                          ' fila 1 = encabezado "Categoría"
        v = UCase$(Trim$(CStr(ws.Cells(i, "A").Value)))
        If Len(v) > 0 Then d(v) = True
    Next i
    Set CargarCategorias = d
End Function

' ============================================================
'  COLAPSAR / EXPANDIR (usado por botones y por la generación de PDF)
' ============================================================
Public Sub MostrarTodo()
    Dim ws As Worksheet, ultima As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Application.ScreenUpdating = False
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=8
    On Error GoTo 0
    ws.Rows(FILA_INICIO & ":" & ultima).EntireRow.Hidden = False
    Application.ScreenUpdating = True
End Sub

Public Sub OcultarDetalle()
    Dim ws As Worksheet, ultima As Long, i As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    MostrarTodo
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Application.ScreenUpdating = False
    For i = FILA_INICIO To ultima
        If EsDetalle(ws, i) Then ws.Rows(i).EntireRow.Hidden = True
    Next i
    Application.ScreenUpdating = True
End Sub

' Restablece la vista a como se abrió el libro originalmente:
'   - Muestra TODAS las filas (deshace cualquier colapso).
'   - Quita por completo el esquema/agrupaciones (los botones [-]/[+]).
'   - Deja la columna marcadora (O) oculta, porque no es contenido visible.
' NO borra las marcas de detalle (los "D" en la columna O se conservan).
Public Sub ResetMacroDefault()
    Dim ws As Worksheet, ultima As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' 1) Expandir cualquier nivel de esquema y mostrar todas las filas.
    On Error Resume Next
    ws.Outline.ShowLevels RowLevels:=8
    On Error GoTo 0
    ws.Rows(FILA_INICIO & ":" & ultima).EntireRow.Hidden = False

    ' 2) Quitar por completo el esquema (deja de haber botones [-]/[+]).
    On Error Resume Next
    ws.Rows.ClearOutline
    On Error GoTo 0

    ' 3) La columna marcadora permanece oculta.
    ws.Columns(COL_MARCA).Hidden = True

    ' 4) Regresar la vista al inicio de la hoja.
    ws.Activate
    Application.Goto ws.Range("A1"), True

    Application.ScreenUpdating = True
    MsgBox "Vista restablecida: todo visible, sin colapsos ni agrupaciones.", vbInformation
End Sub

' ============================================================
'  MARCADO DE FILAS DE DETALLE (columna oculta O = "D")
' ============================================================
Public Sub MarcarComoDetalle()
    MarcarSeleccion MARCA_DETALLE
End Sub

Public Sub QuitarMarcaDetalle()
    MarcarSeleccion ""
End Sub

Private Sub MarcarSeleccion(ByVal valor As String)
    Dim ws As Worksheet, celda As Range, r As Long
    Dim filas As Object
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Selecciona las filas en la hoja de Cotización.", vbExclamation
        Exit Sub
    End If

    Set filas = CreateObject("Scripting.Dictionary")
    For Each celda In Selection.Cells
        r = celda.Row
        If r >= FILA_INICIO And Not filas.Exists(r) Then
            filas(r) = True
            ws.Cells(r, COL_MARCA).Value = valor
        End If
    Next celda
    ws.Columns(COL_MARCA).Hidden = True
    ResaltarDetalle          ' asegura el resaltado rojizo de las filas marcadas
End Sub

' ============================================================
'  RESALTADO ROJIZO DE LAS FILAS MARCADAS (formato condicional)
'  Se pinta solo donde O = "D". No toca el formato original de las celdas;
'  al quitar la marca, el color desaparece automáticamente.
' ============================================================
Public Sub ResaltarDetalle()
    Dim ws As Worksheet, rng As Range, fc As FormatCondition
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    Set rng = RangoDatos(ws)
    If rng Is Nothing Then Exit Sub

    QuitarResaltoDetalle          ' evita duplicar MI formato condicional

    Set fc = rng.FormatConditions.Add(Type:=xlExpression, Formula1:=FormulaResalto())
    fc.Interior.Color = RGB(244, 199, 195)   ' rojizo claro
    fc.StopIfTrue = False
    ws.Columns(COL_MARCA).Hidden = True
End Sub

Public Sub QuitarResaltoDetalle()
    Dim ws As Worksheet, rng As Range, k As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    Set rng = RangoDatos(ws)
    If rng Is Nothing Then Exit Sub

    On Error Resume Next
    For k = rng.FormatConditions.Count To 1 Step -1
        If rng.FormatConditions(k).Type = xlExpression Then
            If rng.FormatConditions(k).Formula1 = FormulaResalto() Then
                rng.FormatConditions(k).Delete
            End If
        End If
    Next k
    On Error GoTo 0
End Sub

Private Function RangoDatos(ByVal ws As Worksheet) As Range
    Dim ultima As Long
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultima < FILA_INICIO Then Exit Function
    Set RangoDatos = ws.Range("A" & FILA_INICIO & ":N" & ultima)
End Function

Private Function FormulaResalto() As String
    ' Ej.: =$O7="D"   (columna fija O, fila relativa desde FILA_INICIO)
    FormulaResalto = "=$" & COL_MARCA & FILA_INICIO & "=""" & MARCA_DETALLE & """"
End Function
