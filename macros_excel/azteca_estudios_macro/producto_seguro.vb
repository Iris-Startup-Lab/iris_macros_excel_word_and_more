' MÓDULO: M_Producto
' Propósito: insertar filas de producto sin pegar formatos de fila completa,
' evitando el error de PasteSpecial en filas con combinaciones distintas.
'
' Depende de:
'   HojaCotizacion(), FILA_INICIO (M_UUID)
'   ValidarFilaEnBloque(), FijarBotones(), DesbloquearBotones(),
'   BloquearBotones(), RestaurarPosicionesBotones() (M_Blindaje)
'   suma (Módulo9)

Option Explicit

Public Sub AgregarProducto()
    Dim ws As Worksheet
    Dim filaBase As Long, filaNueva As Long, filaModelo As Long
    Dim msgError As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Producto"
        Exit Sub
    End If

    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Selecciona una fila dentro de la hoja de Cotización.", vbExclamation, "M_Producto"
        Exit Sub
    End If

    filaBase = ActiveCell.Row
    If Not EsFilaValidaParaProducto(ws, filaBase) Then
        MsgBox "Selecciona una fila de producto o de fecha dentro del bloque " & _
               "para agregar el producto.", vbExclamation, "M_Producto"
        Exit Sub
    End If

    filaModelo = FilaModeloProducto(ws, filaBase)
    If filaModelo = 0 Then
        MsgBox "No se encontró una fila de producto para tomar el formato y las fórmulas.", _
               vbExclamation, "M_Producto"
        Exit Sub
    End If

    FijarBotones
    DesbloquearBotones ws
    Application.ScreenUpdating = False
    On Error GoTo salir

    Rows(filaBase + 1).Insert Shift:=xlDown
    filaNueva = filaBase + 1
    If filaModelo >= filaNueva Then filaModelo = filaModelo + 1

    ' Restaurar combinación A:F, alto, wrap y alineación de la fila modelo.
    PrepararFilaNueva ws, filaModelo, filaNueva

    ' Copiar las fórmulas de K y M sin usar PasteSpecial.
    CopiarFormulasProducto ws, filaModelo, filaNueva

    RestaurarPosicionesBotones
    Call suma

    Application.ScreenUpdating = True
    Exit Sub

salir:
    msgError = Err.Description
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    On Error Resume Next
    RestaurarPosicionesBotones
    BloquearBotones ws
    On Error GoTo 0
    If Len(msgError) > 0 Then
        MsgBox "No se pudo agregar el producto en la fila " & filaNueva & ": " & msgError, _
               vbCritical, "M_Producto"
    End If
End Sub

Public Sub PrepararFilaNueva(ByVal ws As Worksheet, ByVal filaBase As Long, ByVal filaNueva As Long)
    Dim base As Range
    Dim alineacion As Long
    Dim vertical As Long
    Dim wrap As Boolean

    If ws Is Nothing Then Exit Sub

    ' Restaurar las combinaciones horizontales de la fila activa (A:F, M:N, etc.).
    AsegurarCombinacionesFila ws, filaBase, filaNueva

    Set base = ws.Cells(filaBase, "A")
    alineacion = base.HorizontalAlignment
    vertical = base.VerticalAlignment
    wrap = base.WrapText

    On Error Resume Next
    ws.Rows(filaNueva).RowHeight = ws.Rows(filaBase).RowHeight
    With ws.Cells(filaNueva, "A")
        .WrapText = wrap
        .HorizontalAlignment = alineacion
        .VerticalAlignment = vertical
        .IndentLevel = base.IndentLevel
        .NumberFormat = base.NumberFormat
        .Font.Name = base.Font.Name
        .Font.Size = base.Font.Size
        .Font.Bold = base.Font.Bold
        .Font.Color = base.Font.Color
        .Interior.Color = base.Interior.Color
    End With
    Err.Clear
    On Error GoTo 0

    On Error Resume Next
    ws.Cells(filaNueva, "A").Validation.Delete
    Err.Clear
    On Error GoTo 0

    CopiarFormatoColumnasProducto ws, filaBase, filaNueva

    ' Aplicar el estilo de bordes del bloque de productos.
    AplicarBordesProducto ws, filaNueva
End Sub

Private Sub CopiarFormatoColumnasProducto(ByVal ws As Worksheet, ByVal filaModelo As Long, ByVal filaNueva As Long)
    Dim c As Long
    Dim origen As Range, destino As Range

    If ws Is Nothing Then Exit Sub

    For c = 7 To 13
        Set origen = ws.Cells(filaModelo, c)
        Set destino = ws.Cells(filaNueva, c)

        On Error Resume Next
        destino.NumberFormat = origen.NumberFormat
        destino.Font.Name = origen.Font.Name
        destino.Font.Size = origen.Font.Size
        destino.Font.Bold = origen.Font.Bold
        destino.Font.Italic = origen.Font.Italic
        destino.Font.Color = origen.Font.Color
        destino.Interior.Color = origen.Interior.Color
        destino.HorizontalAlignment = origen.HorizontalAlignment
        destino.VerticalAlignment = origen.VerticalAlignment
        destino.WrapText = origen.WrapText
        destino.IndentLevel = origen.IndentLevel
        Err.Clear
        On Error GoTo 0
    Next c
End Sub

Private Sub AsegurarCombinacionesFila(ByVal ws As Worksheet, ByVal filaBase As Long, ByVal filaNueva As Long)
    Dim c As Long
    Dim ma As Range

    c = 1
    Do While c <= 15
        Set ma = ws.Cells(filaBase, c).MergeArea
        If ma.Row = filaBase And ma.Rows.Count = 1 And ma.Columns.Count > 1 Then
            On Error Resume Next
            ws.Range(ws.Cells(filaNueva, ma.Column), _
                     ws.Cells(filaNueva, ma.Column + ma.Columns.Count - 1)).Merge
            Err.Clear
            On Error GoTo 0
            c = ma.Column + ma.Columns.Count
        Else
            c = c + 1
        End If
    Loop
End Sub

Private Sub AplicarBordesProducto(ByVal ws As Worksheet, ByVal fila As Long)
    Dim c As Long
    Dim colorBorde As Long

    colorBorde = RGB(173, 173, 173)

    ' Limpiar cualquier borde heredado o diagonal en la fila nueva.
    On Error Resume Next
    ws.Range(ws.Cells(fila, "A"), ws.Cells(fila, "O")).Borders.LineStyle = xlNone
    Err.Clear
    On Error GoTo 0

    ' Separadores verticales de G a L.
    For c = 7 To 12
        With ws.Cells(fila, c).Borders
            .Item(xlEdgeLeft).LineStyle = xlContinuous
            .Item(xlEdgeLeft).Weight = xlThin
            .Item(xlEdgeLeft).Color = colorBorde

            .Item(xlEdgeRight).LineStyle = xlContinuous
            .Item(xlEdgeRight).Weight = xlThin
            .Item(xlEdgeRight).Color = colorBorde
        End With
    Next c

    ' Separador izquierdo de M (M:N está combinado).
    With ws.Cells(fila, "M").Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = colorBorde
    End With
End Sub

Public Sub CopiarFormulasProducto(ByVal ws As Worksheet, ByVal filaBase As Long, ByVal filaNueva As Long)
    If ws Is Nothing Then Exit Sub

    If ws.Cells(filaBase, "K").HasFormula Then
        ws.Cells(filaNueva, "K").FormulaR1C1 = ws.Cells(filaBase, "K").FormulaR1C1
    End If

    If ws.Cells(filaBase, "M").HasFormula Then
        ws.Cells(filaNueva, "M").FormulaR1C1 = ws.Cells(filaBase, "M").FormulaR1C1
    End If
End Sub

Public Function EsFilaFecha(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim v As Variant
    Dim nf As String

    If ws Is Nothing Then Exit Function

    v = ws.Cells(fila, "A").Value
    If IsDate(v) Then
        EsFilaFecha = True
        Exit Function
    End If

    nf = LCase$(CStr(ws.Cells(fila, "A").NumberFormat))
    If InStr(1, nf, "dd", vbTextCompare) > 0 Or _
       InStr(1, nf, "mm", vbTextCompare) > 0 Or _
       InStr(1, nf, "aaaa", vbTextCompare) > 0 Or _
       InStr(1, nf, "yyyy", vbTextCompare) > 0 Then
        If Len(Trim$(CStr(v))) = 0 Or IsNumeric(v) Then EsFilaFecha = True
    End If
End Function

Public Function VersionMProducto() As String
    VersionMProducto = "2026-11-07-FIX-VALIDACION"
End Function

Public Function FilaModeloProducto(ByVal ws As Worksheet, ByVal filaBase As Long) As Long
    Dim r As Long
    Dim ultima As Long

    If ws Is Nothing Then Exit Function

    If EsFilaProductoModelo(ws, filaBase) Then
        FilaModeloProducto = filaBase
        Exit Function
    End If

    For r = filaBase - 1 To FILA_INICIO Step -1
        If EsFilaProductoModelo(ws, r) Then
            FilaModeloProducto = r
            Exit Function
        End If
        If EsFilaCorte(ws, r) Then Exit For
    Next r

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = filaBase + 1 To ultima
        If EsFilaProductoModelo(ws, r) Then
            FilaModeloProducto = r
            Exit Function
        End If
        If EsFilaCorte(ws, r) Then Exit For
    Next r

    ' Si el bloque quedó sin productos, usar como modelo cualquier producto
    ' de la hoja para no perder formatos ni fórmulas al insertar.
    For r = FILA_INICIO To ultima
        If EsFilaProductoModelo(ws, r) Then
            FilaModeloProducto = r
            Exit Function
        End If
    Next r
End Function

Private Function EsFilaProductoModelo(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim nf As String

    If ws Is Nothing Then Exit Function
    If fila < FILA_INICIO Then Exit Function
    If EsFilaCorte(ws, fila) Then Exit Function
    If ws.Cells(fila, "K").HasFormula Or ws.Cells(fila, "M").HasFormula Then
        nf = CStr(ws.Cells(fila, "G").NumberFormat)
        If InStr(1, nf, "#,##0") > 0 Or InStr(1, nf, "$") > 0 Then
            EsFilaProductoModelo = True
        End If
    End If
End Function

Private Function EsFilaCorte(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim a As String
    Dim ma As Range

    If ws Is Nothing Then Exit Function

    Set ma = ws.Cells(fila, "A").MergeArea
    If ma.Rows.Count >= 2 And ma.Columns.Count >= 6 Then
        EsFilaCorte = True
        Exit Function
    End If

    a = UCase$(Trim$(CStr(ws.Cells(fila, "A").Value)))
    If InStr(1, a, "SUMA TOTAL") > 0 Then EsFilaCorte = True
    If InStr(1, a, "SUBTOTAL") > 0 Then EsFilaCorte = True
    If InStr(1, a, "DESCRIP") > 0 Then EsFilaCorte = True

    a = UCase$(Trim$(CStr(ws.Cells(fila, "I").Value)))
    If InStr(1, a, "SUBTOTAL") > 0 Then EsFilaCorte = True
End Function

Private Function EsFilaValidaParaProducto(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    Dim a As String
    Dim ma As Range

    If ws Is Nothing Then Exit Function
    If fila < FILA_INICIO Then Exit Function
    If Not ValidarFilaEnBloque(fila, False) Then Exit Function

    ' Encabezado de categoría: A:N fusionado en 2 o más filas.
    Set ma = ws.Cells(fila, "A").MergeArea
    If ma.Rows.Count >= 2 And ma.Columns.Count >= 6 Then Exit Function

    a = UCase$(Trim$(CStr(ws.Cells(fila, "A").Value)))
    If InStr(1, a, "SUMA TOTAL") > 0 Then Exit Function
    If InStr(1, a, "DESCRIP") > 0 Or InStr(1, a, "PRECIO") > 0 Or _
       InStr(1, a, "CANTIDAD") > 0 Then Exit Function

    If ws.Cells(fila, "K").HasFormula Or ws.Cells(fila, "M").HasFormula Then
        EsFilaValidaParaProducto = True
        Exit Function
    End If

    If EsFilaFecha(ws, fila) Then
        If FilaModeloProducto(ws, fila) > 0 Then EsFilaValidaParaProducto = True
    End If
End Function

Public Sub AsignarBotonProducto()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim encontrado As Boolean
    Dim msgError As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Producto"
        Exit Sub
    End If

    ws.Activate
    DesbloquearBotones ws
    On Error GoTo salir

    For Each shp In ws.Shapes
        If shp.Type = 8 Then
            If StrComp(shp.Name, "Button 1", vbTextCompare) = 0 Or _
               InStr(1, CStr(shp.AlternativeText), "Agregar producto", vbTextCompare) > 0 Or _
               InStr(1, CStr(shp.OnAction), "AgregarFilaProducto", vbTextCompare) > 0 Then
                shp.OnAction = "AgregarProducto"
                encontrado = True
            End If
        End If
    Next shp

    If Not encontrado Then
        MsgBox "No se encontró el botón 'Agregar producto'.", vbExclamation, "M_Producto"
        GoTo salir
    End If

    FijarBotones ws
    BloquearBotones ws

    MsgBox "Botón 'Agregar producto' asignado a AgregarProducto.", vbInformation, "M_Producto"
    Exit Sub

salir:
    msgError = Err.Description
    On Error Resume Next
    BloquearBotones ws
    On Error GoTo 0
    If Len(msgError) > 0 Then
        MsgBox "No se pudo asignar el botón: " & msgError, vbCritical, "M_Producto"
    End If
End Sub
