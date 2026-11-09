Attribute VB_Name = "M_Catalogos"
' MÓDULO: M_Catalogos
' Propósito:
'   - Catálogo de Categorías: combo temporal a la izquierda del encabezado.
'   - Catálogo de Foros/Espacios: inserta fila y abre el combo en la fila nueva.
'   - Restaura la validación nativa de categorías y foros.
'
' Depende de:
'   HojaCotizacion(), HojaQueContiene(), FILA_INICIO (M_UUID)
'   ValidarFilaEnBloque(), FijarBotones(), DesbloquearBotones(),
'   BloquearBotones(), RestaurarPosicionesBotones() (M_Blindaje)
'   suma (Módulo9)

Option Explicit

Private Const TMP_CATEGORIA As String = "TMP_CatalogoCategoria"
Private Const TMP_FORO As String = "TMP_CatalogoForo"
Private Const PREFIJO_TMP As String = "TMP_"
Private Const TIPO_CATEGORIA As String = "CAT"
Private Const TIPO_FORO As String = "FORO"

Private gFilaCatalogo As Long
Private gTipoCatalogo As String
Private gCatalogoAbierto As Boolean

Public Function VersionMCatalogos() As String
    VersionMCatalogos = "2026-11-07-FIX-MODELO-SHIFT"
End Function

' ============================================================
'  PREPARACIÓN / VALIDACIONES NATIVAS
' ============================================================
Public Sub PrepararCatalogoCategorias()
    Dim ws As Worksheet
    Dim wsCat As Worksheet
    Dim r As Long, ultima As Long, ultimaCat As Long
    Dim formula As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    Set wsCat = HojaQueContiene("CATEGOR")
    If wsCat Is Nothing Then
        MsgBox "No se encontró la hoja de Categorías.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    ultimaCat = wsCat.Cells(wsCat.Rows.Count, "A").End(xlUp).Row
    If ultimaCat < 2 Then
        MsgBox "La hoja de Categorías está vacía.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    On Error GoTo salir

    LimpiarCatalogoTemporal
    DesbloquearBotones ws

    formula = "='" & wsCat.Name & "'!$A$2:$A$" & ultimaCat
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For r = FILA_INICIO To ultima
        If EsFilaEncabezado(ws, r) Then
            InstalarValidacionCategoria ws, r, formula
        End If
    Next r

    BloquearBotones ws
    Application.ScreenUpdating = True

    MsgBox "Validación de categorías restaurada." & vbCrLf & vbCrLf & _
           "Puedes escribir en el encabezado y usar la flecha de la lista." & vbCrLf & _
           "Doble clic en el encabezado abre además el catálogo a la izquierda.", _
           vbInformation, "M_Catalogos"
    Exit Sub

salir:
    BloquearBotones ws
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "Error al preparar el catálogo: " & Err.Description, vbCritical, "M_Catalogos"
    End If
End Sub

Private Sub InstalarValidacionCategoria(ByVal ws As Worksheet, ByVal fila As Long, ByVal formula As String)
    Dim celda As Range

    Set celda = ws.Cells(fila, "A")

    On Error Resume Next
    celda.Validation.Delete
    Err.Clear

    celda.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=formula
    With celda.Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = True
        .ErrorTitle = "Categoría no válida"
        .ErrorMessage = "Selecciona una categoría de la lista."
    End With
    Err.Clear
    On Error GoTo 0
End Sub

' ============================================================
'  CATÁLOGO DE CATEGORÍAS
' ============================================================
'  CATÁLOGO DE CATEGORÍAS
' ============================================================
Public Sub AbrirCatalogoCategoria()
    Dim ws As Worksheet

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Selecciona una celda en la hoja de Cotización.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    AbrirCatalogoCategoriaEn ActiveCell
End Sub

Public Sub AbrirCatalogoCategoriaEn(ByVal Target As Range)
    Dim ws As Worksheet
    Dim fila As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub

    fila = FilaEncabezadoDe(Target)
    If fila = 0 Then Exit Sub

    MostrarComboCategoria ws, fila
End Sub

Public Function EsEncabezadoCategoriaEn(ByVal Target As Range) As Boolean
    EsEncabezadoCategoriaEn = (FilaEncabezadoDe(Target) > 0)
End Function

Public Function EsFilaForoEn(ByVal Target As Range) As Boolean
    Dim ws As Worksheet
    Dim celda As Range
    Dim valor As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Function
    If Target Is Nothing Then Exit Function
    If Target.Worksheet.Name <> ws.Name Then Exit Function

    Set celda = Target.Cells(1, 1)
    If celda.MergeArea.Rows.Count = 1 And celda.MergeArea.Columns.Count >= 6 Then
        Set celda = celda.MergeArea.Cells(1, 1)
    End If

    If celda.Column <> 1 Then Exit Function
    If celda.Row < FILA_INICIO Then Exit Function

    valor = Trim$(CStr(celda.Value))
    If Len(valor) = 0 Then Exit Function
    If InStr(1, valor, " - ", vbTextCompare) > 0 Then
        valor = Trim$(Left$(valor, InStr(1, valor, " - ", vbTextCompare) - 1))
    End If

    EsFilaForoEn = (UCase$(Left$(valor, 5)) = "FORO ")
End Function

Public Sub AbrirCatalogoForoEn(ByVal Target As Range)
    Dim ws As Worksheet
    Dim celda As Range

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub

    Set celda = Target.Cells(1, 1)
    If celda.MergeArea.Rows.Count = 1 And celda.MergeArea.Columns.Count >= 6 Then
        Set celda = celda.MergeArea.Cells(1, 1)
    End If

    If celda.Column <> 1 Then Exit Sub
    If celda.Row < FILA_INICIO Then Exit Sub

    MostrarComboForo ws, celda.Row
End Sub

Private Sub MostrarComboCategoria(ByVal ws As Worksheet, ByVal fila As Long)
    Dim wsCat As Worksheet
    Dim ultima As Long
    Dim rango As String
    Dim celda As Range
    Dim shp As Shape

    Set wsCat = HojaQueContiene("CATEGOR")
    If wsCat Is Nothing Then
        MsgBox "No se encontró la hoja de Categorías.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    ultima = wsCat.Cells(wsCat.Rows.Count, "A").End(xlUp).Row
    If ultima < 2 Then
        MsgBox "La hoja de Categorías está vacía.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    LimpiarCatalogoTemporal
    DesbloquearBotones ws

    rango = "'" & wsCat.Name & "'!$A$2:$A$" & ultima
    Set celda = ws.Cells(fila, "A")
    Set shp = CrearComboCatalogo(ws, celda, rango, TMP_CATEGORIA, "AplicarCategoriaCatalogo", 22, 160)
    If shp Is Nothing Then
        BloquearBotones ws
        Exit Sub
    End If

    gFilaCatalogo = fila
    gTipoCatalogo = TIPO_CATEGORIA
    gCatalogoAbierto = True

    On Error Resume Next
    If Len(Trim$(CStr(ws.Cells(fila, "A").Value))) > 0 Then
        shp.ControlFormat.Value = ws.Cells(fila, "A").Value
    End If
    shp.ControlFormat.DropDown
    On Error GoTo 0
End Sub

Public Sub AplicarCategoriaCatalogo()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim valor As String
    Dim idx As Long

    If gTipoCatalogo <> TIPO_CATEGORIA Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    Set shp = BuscarComboTemporal(ws)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    idx = shp.ControlFormat.ListIndex
    If idx >= 1 Then
        valor = CStr(shp.ControlFormat.List(idx))
    Else
        valor = CStr(shp.ControlFormat.Value)
    End If
    On Error GoTo 0

    If Len(Trim$(valor)) > 0 And gFilaCatalogo >= FILA_INICIO Then
        ws.Cells(gFilaCatalogo, "A").Value = valor
    End If

    LimpiarCatalogoTemporal
    RestaurarPosicionesBotones
End Sub

' ============================================================
'  CATÁLOGO DE FOROS / ESPACIOS
' ============================================================
Public Sub AgregarEspacioForo()
    Dim ws As Worksheet
    Dim filaBase As Long, filaNueva As Long, filaModelo As Long
    Dim msgError As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    If ActiveSheet.Name <> ws.Name Then
        MsgBox "Selecciona una fila dentro de la hoja de Cotización.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    RestaurarVistaCotizacion
    filaBase = ActiveCell.Row
    If Not EsFilaValidaParaForo(ws, filaBase) Then
        MsgBox "Selecciona una fila de producto o de fecha dentro del bloque " & _
               "para agregar el espacio/foro.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    filaModelo = FilaModeloProducto(ws, filaBase)
    If filaModelo = 0 Then
        MsgBox "No se encontró una fila de producto para tomar el formato y las fórmulas.", _
               vbExclamation, "M_Catalogos"
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

    ' Desplazar solo si la fila nueva quedó fuera de la ventana visible.
    On Error Resume Next
    If filaNueva < ActiveWindow.VisibleRange.Row Or _
       filaNueva > ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count - 1 Then
        Application.Goto ws.Cells(filaNueva, "A"), True
    End If
    On Error GoTo 0

    MostrarComboForo ws, filaNueva
    If Not gCatalogoAbierto Then BloquearBotones ws
    Application.ScreenUpdating = True
    Exit Sub

salir:
    msgError = Err.Description
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    On Error Resume Next
    LimpiarCatalogoTemporal
    BloquearBotones ws
    On Error GoTo 0
    If Len(msgError) > 0 Then
        MsgBox "No se pudo agregar el espacio/foro en la fila " & filaNueva & ": " & msgError, _
               vbCritical, "M_Catalogos"
    End If
End Sub

Private Function EsFilaValidaParaForo(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
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

    ' Debe tener fórmulas de producto en K o M, o ser una fila de fecha
    ' con una fila modelo de producto disponible.
    If ws.Cells(fila, "K").HasFormula Or ws.Cells(fila, "M").HasFormula Then
        EsFilaValidaParaForo = True
        Exit Function
    End If

    If EsFilaFecha(ws, fila) Then
        If FilaModeloProducto(ws, fila) > 0 Then EsFilaValidaParaForo = True
    End If
End Function

Private Sub MostrarComboForo(ByVal ws As Worksheet, ByVal fila As Long)
    Dim wsEsp As Worksheet
    Dim ultima As Long
    Dim rango As String
    Dim celda As Range
    Dim shp As Shape
    Dim altoFila As Double
    Dim anchoCombo As Double
    Dim actual As String

    Set wsEsp = HojaQueContiene("ESPACI")
    If wsEsp Is Nothing Then
        MsgBox "No se encontró la hoja Espacios.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    ultima = wsEsp.Cells(wsEsp.Rows.Count, "A").End(xlUp).Row
    If ultima < 2 Then
        MsgBox "La hoja Espacios está vacía.", vbExclamation, "M_Catalogos"
        Exit Sub
    End If

    LimpiarCatalogoTemporal
    DesbloquearBotones ws

    rango = "'" & wsEsp.Name & "'!$A$2:$A$" & ultima
    Set celda = ws.Cells(fila, "A")
    altoFila = ws.Rows(fila).RowHeight
    If altoFila < 14 Then altoFila = 14
    anchoCombo = celda.MergeArea.Width - 4
    If anchoCombo < 120 Then anchoCombo = 120

    ' Mientras el combo temporal está abierto no debe aparecer también la
    ' flecha de la validación nativa: se restaura al cerrar el combo.
    On Error Resume Next
    celda.Validation.Delete
    On Error GoTo 0

    Set shp = CrearComboCatalogo(ws, celda, rango, TMP_FORO, "AplicarForoCatalogo", altoFila, anchoCombo)
    If shp Is Nothing Then
        BloquearBotones ws
        Exit Sub
    End If

    gFilaCatalogo = fila
    gTipoCatalogo = TIPO_FORO
    gCatalogoAbierto = True

    On Error Resume Next
    actual = Trim$(CStr(ws.Cells(fila, "A").Value))
    If InStr(1, actual, " - ", vbTextCompare) > 0 Then
        actual = Trim$(Left$(actual, InStr(1, actual, " - ", vbTextCompare) - 1))
    End If
    If Len(actual) > 0 Then shp.ControlFormat.Value = actual
    Err.Clear
    shp.ControlFormat.DropDown
    On Error GoTo 0
End Sub

Public Sub AplicarForoCatalogo()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim valor As String
    Dim descripcion As String
    Dim idx As Long

    If gTipoCatalogo <> TIPO_FORO Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    Set shp = BuscarComboTemporal(ws)
    If shp Is Nothing Then Exit Sub

    On Error Resume Next
    idx = shp.ControlFormat.ListIndex
    If idx >= 1 Then
        valor = CStr(shp.ControlFormat.List(idx))
    Else
        valor = CStr(shp.ControlFormat.Value)
    End If
    On Error GoTo 0

    If Len(Trim$(valor)) > 0 And gFilaCatalogo >= FILA_INICIO Then
        descripcion = DescripcionEspacio(valor)
        If Len(descripcion) > 0 Then
            ws.Cells(gFilaCatalogo, "A").Value = valor & " - " & descripcion
        Else
            ws.Cells(gFilaCatalogo, "A").Value = valor
        End If
    End If

    LimpiarCatalogoTemporal
    Call suma
    RestaurarPosicionesBotones
    RestaurarVistaCotizacion
End Sub

Public Sub RestaurarVistaCotizacion()
    On Error Resume Next
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 0
    On Error GoTo 0
End Sub

Private Function DescripcionEspacio(ByVal id As String) As String
    Dim wsEsp As Worksheet
    Dim ultima As Long
    Dim idx As Variant

    Set wsEsp = HojaQueContiene("ESPACI")
    If wsEsp Is Nothing Then Exit Function

    ultima = wsEsp.Cells(wsEsp.Rows.Count, "A").End(xlUp).Row
    If ultima < 2 Then Exit Function

    On Error Resume Next
    idx = Application.Match(id, wsEsp.Range("A2:A" & ultima), 0)
    On Error GoTo 0

    If Not IsError(idx) And IsNumeric(idx) Then
        DescripcionEspacio = Trim$(CStr(wsEsp.Cells(CLng(idx) + 1, "B").Value))
    End If
End Function

Public Sub AjustarDescripcionForo(ByVal Target As Range)
    Dim ws As Worksheet
    Dim celda As Range
    Dim valor As String
    Dim descripcion As String
    Static ocupado As Boolean

    If ocupado Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub

    Set celda = Target.Cells(1, 1)
    If celda.MergeArea.Rows.Count = 1 Then
        Set celda = celda.MergeArea.Cells(1, 1)
    End If

    If celda.Column <> 1 Then Exit Sub
    If celda.Row < FILA_INICIO Then Exit Sub

    valor = Trim$(CStr(celda.Value))
    If Len(valor) = 0 Then Exit Sub
    If InStr(1, valor, " - ", vbTextCompare) > 0 Then Exit Sub

    descripcion = DescripcionEspacio(valor)
    If Len(descripcion) = 0 Then Exit Sub

    ocupado = True
    On Error GoTo salir
    Application.EnableEvents = False
    celda.Value = valor & " - " & descripcion

salir:
    Application.EnableEvents = True
    ocupado = False
End Sub

' ============================================================
'  INSTALADOR DEL BOTÓN DE FOROS
' ============================================================
Public Sub AsignarBotonForo()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim asignado As Boolean
    Dim eliminado As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    DesbloquearBotones ws
    On Error Resume Next

    ' Eliminar el control huérfano confirmado por el usuario
    Set shp = Nothing
    Set shp = ws.Shapes("Drop Down 31")
    If Not shp Is Nothing Then
        If shp.FormControlType = xlDropDown Then
            shp.Delete
            eliminado = True
        End If
    End If

    ' Asignar la macro al botón de foros
    For Each shp In ws.Shapes
        If StrComp(shp.Name, "Button 9", vbTextCompare) = 0 Then
            shp.OnAction = "AgregarEspacioForo"
            asignado = True
        ElseIf InStr(1, CStr(shp.AlternativeText), "espacio/foro", vbTextCompare) > 0 Then
            shp.OnAction = "AgregarEspacioForo"
            asignado = True
        End If
    Next shp

    On Error GoTo 0
    BloquearBotones ws

    MsgBox "Botón de foros asignado: " & asignado & vbCrLf & _
           "Control huérfano eliminado: " & eliminado, _
           vbInformation, "M_Catalogos"
End Sub

' ============================================================
'  LIMPIEZA / ESTADO
' ============================================================
Public Sub QuitarValidacionesForo()
    Dim ws As Worksheet
    Dim r As Long, ultima As Long
    Dim f As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    DesbloquearBotones ws
    On Error Resume Next

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > ultima Then
        ultima = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    End If

    For r = FILA_INICIO To ultima
        f = ""
        f = ws.Cells(r, "A").Validation.Formula1
        If InStr(1, f, "Espacios", vbTextCompare) > 0 Then
            ws.Cells(r, "A").Validation.Delete
        End If
    Next r

    Err.Clear
    On Error GoTo 0
    BloquearBotones ws

    MsgBox "Validaciones de foro eliminadas." & vbCrLf & _
           "El catálogo de foros se abrirá con doble clic en la fila.", _
           vbInformation, "M_Catalogos"
End Sub

Public Sub LimpiarControlesHuerfanos()
    Dim ws As Worksheet
    Dim sh As Shape
    Dim i As Long
    Dim n As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "M_Catalogos"
        Exit Sub
    End If

    DesbloquearBotones ws
    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If InStr(1, sh.Name, PREFIJO_TMP, vbTextCompare) = 1 Then
            sh.Delete
            n = n + 1
        ElseIf sh.FormControlType = xlDropDown Then
            If Len(sh.ControlFormat.ListFillRange) = 0 And Len(sh.OnAction) = 0 Then
                sh.Delete
                n = n + 1
            End If
        End If
    Next i

    On Error GoTo 0
    BloquearBotones ws

    MsgBox n & " control(es) huérfano(s) eliminado(s).", vbInformation, "M_Catalogos"
End Sub

Public Sub LimpiarCatalogoSiCorresponde(ByVal Target As Range)
    Dim ws As Worksheet
    Dim fila As Long

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Not gCatalogoAbierto Then Exit Sub

    If Target Is Nothing Then
        LimpiarCatalogoTemporal
        Exit Sub
    End If

    Select Case gTipoCatalogo
        Case TIPO_CATEGORIA
            fila = FilaEncabezadoDe(Target)
            If fila <> gFilaCatalogo Then LimpiarCatalogoTemporal
        Case TIPO_FORO
            If Target.Cells(1, 1).Row <> gFilaCatalogo Then LimpiarCatalogoTemporal
        Case Else
            LimpiarCatalogoTemporal
    End Select
End Sub

Public Sub LimpiarCatalogoTemporal()
    Dim ws As Worksheet
    Dim i As Long
    Dim habia As Boolean
    Dim estabaAbierto As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    estabaAbierto = gCatalogoAbierto
    If Not estabaAbierto And Not HayComboTemporal(ws) Then Exit Sub

    DesbloquearBotones ws
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(1, ws.Shapes(i).Name, PREFIJO_TMP, vbTextCompare) = 1 Then
            ws.Shapes(i).Delete
            habia = True
        End If
    Next i
    On Error GoTo 0

    gCatalogoAbierto = False
    gFilaCatalogo = 0
    gTipoCatalogo = ""

    If habia Or estabaAbierto Then BloquearBotones ws
End Sub

' ============================================================
'  AUXILIARES
' ============================================================
Private Function CrearComboCatalogo(ByVal ws As Worksheet, ByVal celda As Range, _
                                    ByVal rango As String, ByVal nombre As String, _
                                    ByVal macroOnAction As String, ByVal altoCombo As Double, _
                                    ByVal anchoCombo As Double) As Shape
    Dim shp As Shape
    Dim zoomAnterior As Long
    Dim msgError As String

    zoomAnterior = 100
    On Error GoTo salir

    Application.ScreenUpdating = False
    zoomAnterior = ActiveWindow.Zoom
    ActiveWindow.Zoom = 100

    Set shp = ws.Shapes.AddFormControl(xlDropDown, 0, 0, anchoCombo, altoCombo)
    With shp
        .Name = nombre
        .Placement = xlFreeFloating
        .Locked = True
        .Left = celda.MergeArea.Left + 2
        If .Height > celda.MergeArea.Height Then
            .Top = celda.Top
        Else
            .Top = celda.Top + (celda.MergeArea.Height - .Height) / 2
        End If
        .OnAction = macroOnAction
        .ControlFormat.ListFillRange = rango
        .ControlFormat.DropDownLines = 12
    End With

    ActiveWindow.Zoom = zoomAnterior
    Application.ScreenUpdating = True

    Set CrearComboCatalogo = shp
    Exit Function

salir:
    msgError = Err.Description
    On Error Resume Next
    ActiveWindow.Zoom = zoomAnterior
    Application.ScreenUpdating = True
    If Not shp Is Nothing Then shp.Delete
    Set CrearComboCatalogo = Nothing
    If Len(msgError) > 0 Then
        MsgBox "No se pudo crear el catálogo: " & msgError, vbExclamation, "M_Catalogos"
    End If
End Function

Private Function FilaEncabezadoDe(ByVal Target As Range) As Long
    Dim ws As Worksheet
    Dim celda As Range

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Function
    If Target Is Nothing Then Exit Function
    If Target.Worksheet.Name <> ws.Name Then Exit Function

    Set celda = Target.Cells(1, 1)
    If EsFilaEncabezado(ws, celda.MergeArea.Row) Then
        FilaEncabezadoDe = celda.MergeArea.Row
    End If
End Function

Private Function EsFilaEncabezado(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim celda As Range
    Dim v As String

    If ws Is Nothing Then Exit Function
    If r < FILA_INICIO Then Exit Function

    Set celda = ws.Cells(r, "A")

    ' Encabezado fusionado típico: A7:N8 (2 filas, varias columnas)
    If celda.MergeArea.Rows.Count >= 2 And celda.MergeArea.Columns.Count >= 6 Then
        EsFilaEncabezado = True
        Exit Function
    End If

    ' Respaldo: si el texto coincide con una categoría y no es SUMA
    v = UCase$(Trim$(CStr(celda.Value)))
    If Len(v) > 0 Then
        If InStr(1, v, "SUMA") = 0 Then
            EsFilaEncabezado = EsCategoria(v)
        End If
    End If
End Function

Private Function EsCategoria(ByVal texto As String) As Boolean
    Dim cats As Object

    Set cats = CargarCategorias()
    EsCategoria = cats.Exists(UCase$(Trim$(texto)))
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
    For i = 2 To ultima
        v = UCase$(Trim$(CStr(ws.Cells(i, "A").Value)))
        If Len(v) > 0 Then d(v) = True
    Next i

    Set CargarCategorias = d
End Function

Private Function HayComboTemporal(ByVal ws As Worksheet) As Boolean
    Dim shp As Shape

    On Error Resume Next
    For Each shp In ws.Shapes
        If InStr(1, shp.Name, PREFIJO_TMP, vbTextCompare) = 1 Then
            HayComboTemporal = True
            Exit Function
        End If
    Next shp
    On Error GoTo 0
End Function

Private Function BuscarComboTemporal(ByVal ws As Worksheet) As Shape
    Dim shp As Shape

    On Error Resume Next
    For Each shp In ws.Shapes
        If InStr(1, shp.Name, "TMP_Catalogo", vbTextCompare) = 1 Then
            Set BuscarComboTemporal = shp
            Exit Function
        End If
    Next shp
    On Error GoTo 0
End Function
