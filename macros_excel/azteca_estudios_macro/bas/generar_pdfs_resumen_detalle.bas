Attribute VB_Name = "M_PDF"
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

Public Function VersionMPDF() As String
    VersionMPDF = "2026-11-07-ABRE-DETALLE-FIX"
End Function

Public Sub GenerarPDFsCotizacion()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical, "Error"
        Exit Sub
    End If

    ' Validar que el libro esté guardado en disco
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Por favor, guarda el libro de Excel en tu computadora antes de generar los PDFs." & vbCrLf & _
               "(Archivo > Guardar como...)", vbExclamation, "Libro no guardado"
        Exit Sub
    End If

    ' Eliminar cualquier combo temporal del catálogo antes de exportar
    On Error Resume Next
    DesbloquearBotones ws
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(1, ws.Shapes(i).Name, "TMP_Catalogo", vbTextCompare) > 0 Then
            ws.Shapes(i).Delete
        End If
    Next i
    BloquearBotones ws
    On Error GoTo 0

    ' Garantiza que el UUID (K4) esté asignado y sea consistente con el proyecto
    AsegurarUUIDCotizacion

    ' Recálculo automático: asegura que todos los totales reflejen los datos actuales
    On Error Resume Next
    Call suma
    On Error GoTo 0

    ' Ocultar filas para el Resumen aplasta/mueve los objetos atados a celdas:
    ' los anclamos y los devolvemos a su panel antes de exportar.
    FijarBotones ws
    RestaurarPosicionesBotones

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
        MsgBox "Falta el nombre del proyecto (C4) para poder asignar el ID/UUID.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If

    ruta = ThisWorkbook.Path & "\"
    baseNombre = "Cotizacion_" & cliente & "_" & proyecto & "_" & uuid
    If Len(version) > 0 Then baseNombre = baseNombre & "_v" & version
    If Len(fechaCot) > 0 Then baseNombre = baseNombre & "_" & fechaCot

    Dim ultima As Long
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
    If Not ExportarPDF(ws, ruta & baseNombre & "_Detalle.pdf") Then
        GoTo FinalizarConError
    End If

    ' 2) RESUMEN: se restaura EXACTAMENTE lo que el usuario tenía colapsado.
    For i = FILA_INICIO To ultima
        ws.Rows(i).EntireRow.Hidden = estado(i)
    Next i
    If Not ExportarPDF(ws, ruta & baseNombre & "_Resumen.pdf") Then
        GoTo FinalizarConError
    End If

    ' La hoja queda en el estado del usuario; devolvemos el rojizo de trabajo.
    ResaltarDetalle
    Application.ScreenUpdating = True

    MsgBox "PDFs generados exitosamente en la carpeta del libro:" & vbCrLf & vbCrLf & _
           "  " & baseNombre & "_Detalle.pdf" & vbCrLf & _
           "  " & baseNombre & "_Resumen.pdf" & vbCrLf & vbCrLf & _
           "ID Cotización (UUID): " & uuid & vbCrLf & vbCrLf & _
           "Se abrirá el PDF de Detalle.", vbInformation, "Exportación Dual Exitosa"

    On Error Resume Next
    ThisWorkbook.FollowHyperlink ruta & baseNombre & "_Detalle.pdf"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "explorer.exe /select,""" & ruta & baseNombre & "_Detalle.pdf""", vbNormalFocus
    End If
    On Error GoTo 0
    Exit Sub

FinalizarConError:
    ' Restaurar estado de filas y resalto si ocurrió un error durante la exportación
    For i = FILA_INICIO To ultima
        ws.Rows(i).EntireRow.Hidden = estado(i)
    Next i
    ResaltarDetalle
    Application.ScreenUpdating = True
End Sub

Private Function ExportarPDF(ByVal ws As Worksheet, ByVal archivo As String) As Boolean
    Dim formatoA1 As String
    Dim numError As Long, descError As String

    On Error Resume Next
    ' Ocultar temporalmente el aviso "Esta es una plantilla" en A1 para que no salga en el PDF
    formatoA1 = ws.Range("A1").NumberFormat
    ws.Range("A1").NumberFormat = ";;;"
    Err.Clear

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=archivo, _
        Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    numError = Err.Number
    descError = Err.Description

    ' Restaurar siempre la visibilidad del texto en A1
    ws.Range("A1").NumberFormat = formatoA1
    On Error GoTo 0

    If numError <> 0 Then
        If numError = 1004 Then
            MsgBox "No se pudo generar el archivo PDF porque ya se encuentra abierto en otro programa " & _
                   "(como Adobe Acrobat o su navegador web):" & vbCrLf & vbCrLf & _
                   "  " & archivo & vbCrLf & vbCrLf & _
                   "Por favor ciérralo y vuelve a presionar el botón.", vbCritical, "Archivo PDF abierto o bloqueado"
        Else
            MsgBox "Ocurrió un error inesperado al generar el PDF (" & numError & "):" & vbCrLf & _
                   descError & vbCrLf & vbCrLf & "Archivo: " & archivo, vbCritical, "Error al exportar PDF"
        End If
        ExportarPDF = False
        Exit Function
    End If
    ExportarPDF = True
End Function

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

