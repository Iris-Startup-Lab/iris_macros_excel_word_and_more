' Autor: Fernando Dorantes Nieto
'
' MÓDULO: M_PDF
' Propósito: con un solo clic generar DOS PDFs de la cotización:
'   1) *_Detalle.pdf  -> todo visible (el cliente ve el detalle completo).
'   2) *_Resumen.pdf  -> se ocultan las filas de detalle marcadas ("D" en col O).
'
' El nombre de archivo incluye el UUID (ID Cotización), de modo que archivo e
' identificador interno siempre coinciden. Los totales (SUMA TOTAL, SUBTOTAL,
' IVA, TOTAL) NO cambian al colapsar: son valores ya calculados, así que el
' Resumen muestra menos líneas pero el mismo total.
'
' Depende de: HojaCotizacion(), AsegurarUUIDCotizacion() (M_UUID)
'             MostrarTodo(), OcultarDetalle() (M_Esquema).

Option Explicit

Public Sub GenerarPDFsCotizacion()
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
        Exit Sub
    End If

    ' Garantiza que el UUID (K4) esté asignado y sea consistente con el proyecto.
    AsegurarUUIDCotizacion

    Dim cliente As String, proyecto As String, uuid As String
    Dim version As String, fechaCot As String
    Dim ruta As String, baseNombre As String

    cliente = LimpiarNombre(CStr(ws.Range("C3").Value))
    proyecto = LimpiarNombre(CStr(ws.Range("C4").Value))
    uuid = Trim$(CStr(ws.Range("K4").Value))
    version = LimpiarNombre(CStr(ws.Range("N4").Value))
    On Error Resume Next
    fechaCot = Format(ws.Range("K3").Value, "dd-mm-yyyy")
    On Error GoTo 0

    If Len(uuid) = 0 Then
        MsgBox "Falta el nombre del proyecto (C4) para poder asignar el ID/UUID.", vbExclamation
        Exit Sub
    End If

    ruta = ThisWorkbook.Path & "\"
    baseNombre = "Cotizacion_" & cliente & "_" & proyecto & "_" & uuid
    If Len(version) > 0 Then baseNombre = baseNombre & "_v" & version
    If Len(fechaCot) > 0 Then baseNombre = baseNombre & "_" & fechaCot

    Dim ultima As Long, i As Long
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Guardamos el estado de colapso ACTUAL del usuario (lo que dejó oculto a
    ' mano o con los botones), para que el RESUMEN lo respete tal cual.
    Dim estado() As Boolean
    ReDim estado(FILA_INICIO To ultima)
    For i = FILA_INICIO To ultima
        estado(i) = ws.Rows(i).EntireRow.Hidden
    Next i

    Application.ScreenUpdating = False
    QuitarResaltoDetalle          ' PDFs limpios: sin el rojizo de trabajo

    ' 1) DETALLE: todo visible (se expande todo temporalmente).
    MostrarTodo
    Exportar ws, ruta & baseNombre & "_Detalle.pdf"

    ' 2) RESUMEN: se restaura EXACTAMENTE lo que el usuario tenía colapsado.
    For i = FILA_INICIO To ultima
        ws.Rows(i).EntireRow.Hidden = estado(i)
    Next i
    Exportar ws, ruta & baseNombre & "_Resumen.pdf"

    ' La hoja queda en el estado del usuario; devolvemos el rojizo de trabajo.
    ResaltarDetalle

    Application.ScreenUpdating = True

    MsgBox "PDFs generados en la carpeta del libro:" & vbCrLf & vbCrLf & _
           "  - " & baseNombre & "_Detalle.pdf" & vbCrLf & _
           "  - " & baseNombre & "_Resumen.pdf" & vbCrLf & vbCrLf & _
           "ID Cotización (UUID): " & uuid, vbInformation
End Sub

Private Sub Exportar(ByVal ws As Worksheet, ByVal archivo As String)
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=archivo, _
        Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
End Sub

Private Function LimpiarNombre(ByVal s As String) As String
    Dim r As String
    r = Trim$(s)
    r = Replace(r, "/", "-")
    r = Replace(r, "\", "-")
    r = Replace(r, ":", "-")
    r = Replace(r, "*", "")
    r = Replace(r, "?", "")
    r = Replace(r, Chr$(34), "")   ' comillas dobles
    r = Replace(r, "<", "")
    r = Replace(r, ">", "")
    r = Replace(r, "|", "")
    LimpiarNombre = r
End Function
