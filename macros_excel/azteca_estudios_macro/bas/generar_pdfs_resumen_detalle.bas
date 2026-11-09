Attribute VB_Name = "M_PDF"
' Autores: Fernando Dorantes Nieto
' Diana Berumen Estrada
'
' M�DULO: M_PDF
' Prop�sito: con un solo clic generar DOS PDFs de la cotizaci�n:
'   1) *_Detalle.pdf  -> todo visible (el cliente ve el detalle completo).
'   2) *_Resumen.pdf  -> se ocultan las filas de detalle marcadas ("D" en col O).
'
' El nombre de archivo incluye el UUID (ID Cotizaci�n), de modo que archivo e
' identificador interno siempre coinciden. Los totales (SUMA TOTAL, SUBTOTAL,
' IVA, TOTAL) NO cambian al colapsar: son valores ya calculados, as� que el
' Resumen muestra menos l�neas pero el mismo total.
'
' Depende de: HojaCotizacion(), AsegurarUUIDCotizacion() (M_UUID)
'             MostrarTodo(), OcultarDetalle() (M_Esquema).

Option Explicit

Public Function VersionMPDF() As String
    VersionMPDF = "2026-11-09-NOMBRE-SEGURO-FIX"
End Function

Public Sub GenerarPDFsCotizacion()
    Dim ws As Worksheet
    Dim i As Long
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontr� la hoja de Cotizaci�n.", vbCritical, "Error"
        Exit Sub
    End If

    ' Validar que el libro est� guardado en disco
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Por favor, guarda el libro de Excel en tu computadora antes de generar los PDFs." & vbCrLf & _
               "(Archivo > Guardar como...)", vbExclamation, "Libro no guardado"
        Exit Sub
    End If

    ' Eliminar cualquier combo temporal del cat�logo antes de exportar
    On Error Resume Next
    DesbloquearBotones ws
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(1, ws.Shapes(i).Name, "TMP_Catalogo", vbTextCompare) > 0 Then
            ws.Shapes(i).Delete
        End If
    Next i
    BloquearBotones ws
    On Error GoTo 0

    ' Garantiza que el UUID (K4) est� asignado y sea consistente con el proyecto
    AsegurarUUIDCotizacion

    ' Rec�lculo autom�tico: asegura que todos los totales reflejen los datos actuales
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
    ' La fecha puede venir como texto ("10/05/2026"): se higieniza igual que el
    ' resto para que una "/" no rompa la ruta del archivo.
    If IsDate(ws.Range("K3").Value) Then
        fechaCot = Format(CDate(ws.Range("K3").Value), "dd-mm-yyyy")
    Else
        fechaCot = LimpiarNombre(CStr(ws.Range("K3").Value))
    End If

    ' Sin proyecto no hay identidad: se valida C4, no K4 (K4 podr�a traer texto
    ' escrito a mano). Y el ID debe ser exactamente el guardado en Meta.
    If Len(proyecto) = 0 Then
        MsgBox "Falta el nombre del proyecto (C4) para poder asignar el ID/UUID.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    If Len(uuid) = 0 Or StrComp(uuid, UUIDGuardado(), vbTextCompare) <> 0 Then
        MsgBox "El ID Cotizaci�n (K4) no coincide con el registrado para este proyecto." & vbCrLf & _
               "No se generaron los PDFs. Vuelve a escribir el nombre del proyecto en C4.", _
               vbExclamation, "ID Cotizaci�n inv�lido"
        Exit Sub
    End If

    ruta = ThisWorkbook.Path & "\"
    AjustarLargoNombre ruta, cliente, proyecto, uuid, version, fechaCot
    baseNombre = ArmarBaseNombre(cliente, proyecto, uuid, version, fechaCot)

    Dim ultima As Long
    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ultima < FILA_INICIO Then ultima = FILA_INICIO

    ' Guardamos el estado de colapso ACTUAL del usuario (lo que dej� oculto a
    ' mano o con los botones), para que el RESUMEN lo respete tal cual.
    Dim estado() As Boolean
    ReDim estado(FILA_INICIO To ultima)
    For i = FILA_INICIO To ultima
        estado(i) = ws.Rows(i).EntireRow.Hidden
    Next i

    ' A partir de aqu� la hoja se altera (filas ocultas, resalto): cualquier
    ' error inesperado debe regresarla al estado del usuario.
    On Error GoTo ErrorInesperado
    Application.ScreenUpdating = False
    QuitarResaltoDetalle          ' PDFs limpios: sin el rojizo de trabajo

    ' 1) DETALLE: todo visible (se expande todo temporalmente).
    MostrarTodo
    If Not ExportarPDF(ws, ruta & baseNombre & "_Detalle.pdf") Then
        GoTo FinalizarConError
    End If

    ' 2) RESUMEN: se restaura EXACTAMENTE lo que el usuario ten�a colapsado.
    For i = FILA_INICIO To ultima
        ws.Rows(i).EntireRow.Hidden = estado(i)
    Next i
    If Not ExportarPDF(ws, ruta & baseNombre & "_Resumen.pdf") Then
        GoTo FinalizarConError
    End If

    ' La hoja queda en el estado del usuario; devolvemos el rojizo de trabajo.
    On Error GoTo 0
    ResaltarDetalle
    Application.ScreenUpdating = True

    MsgBox "PDFs generados exitosamente en la carpeta del libro:" & vbCrLf & vbCrLf & _
           "  " & baseNombre & "_Detalle.pdf" & vbCrLf & _
           "  " & baseNombre & "_Resumen.pdf" & vbCrLf & vbCrLf & _
           "ID Cotizaci�n (UUID): " & uuid & vbCrLf & vbCrLf & _
           "Se abrir� el PDF de Detalle.", vbInformation, "Exportaci�n Dual Exitosa"

    On Error Resume Next
    ThisWorkbook.FollowHyperlink ruta & baseNombre & "_Detalle.pdf"
    If Err.Number <> 0 Then
        Err.Clear
        Shell "explorer.exe /select,""" & ruta & baseNombre & "_Detalle.pdf""", vbNormalFocus
    End If
    On Error GoTo 0
    Exit Sub

ErrorInesperado:
    MsgBox "Ocurri� un error inesperado al preparar los PDFs (" & Err.Number & "):" & vbCrLf & _
           Err.Description & vbCrLf & vbCrLf & "La hoja se regres� a como estaba.", _
           vbCritical, "Error al generar PDFs"
    Resume FinalizarConError

FinalizarConError:
    ' Restaurar estado de filas y resalto si ocurri� un error durante la exportaci�n
    On Error Resume Next
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
            MsgBox "No se pudo generar el archivo PDF:" & vbCrLf & vbCrLf & _
                   "  " & archivo & vbCrLf & vbCrLf & _
                   "Causas m�s comunes:" & vbCrLf & _
                   "  � El PDF ya est� abierto (Acrobat, Edge, navegador). Ci�rralo." & vbCrLf & _
                   "  � La carpeta es de solo lectura o est� sincroniz�ndose (OneDrive)." & vbCrLf & _
                   "  � No hay una impresora predeterminada configurada en Windows." & vbCrLf & vbCrLf & _
                   "Corrige y vuelve a presionar el bot�n.", vbCritical, "No se pudo exportar el PDF"
        Else
            MsgBox "Ocurri� un error inesperado al generar el PDF (" & numError & "):" & vbCrLf & _
                   descError & vbCrLf & vbCrLf & "Archivo: " & archivo, vbCritical, "Error al exportar PDF"
        End If
        ExportarPDF = False
        Exit Function
    End If
    ExportarPDF = True
End Function

Private Function ArmarBaseNombre(ByVal cliente As String, ByVal proyecto As String, _
                                 ByVal uuid As String, ByVal version As String, _
                                 ByVal fechaCot As String) As String
    Dim s As String
    s = "Cotizacion_" & cliente & "_" & proyecto & "_" & uuid
    If Len(version) > 0 Then s = s & "_v" & version
    If Len(fechaCot) > 0 Then s = s & "_" & fechaCot
    ArmarBaseNombre = s
End Function

' ExportAsFixedFormat falla (1004) con rutas de m�s de ~218 caracteres. Si la
' ruta completa se pasa, se recortan cliente y proyecto (nunca el UUID).
Private Sub AjustarLargoNombre(ByVal ruta As String, ByRef cliente As String, _
                               ByRef proyecto As String, ByVal uuid As String, _
                               ByVal version As String, ByVal fechaCot As String)
    Const LARGO_MAX As Long = 218
    Dim exceso As Long, recorte As Long

    exceso = Len(ruta & ArmarBaseNombre(cliente, proyecto, uuid, version, fechaCot) & "_Resumen.pdf") - LARGO_MAX
    If exceso <= 0 Then Exit Sub

    ' Primero se recorta el proyecto (dejando al menos 10 caracteres)...
    recorte = exceso
    If recorte > Len(proyecto) - 10 Then recorte = Len(proyecto) - 10
    If recorte > 0 Then
        proyecto = Trim$(Left$(proyecto, Len(proyecto) - recorte))
        exceso = exceso - recorte
    End If
    ' ...y luego el cliente, si todav�a hace falta.
    If exceso > 0 Then
        recorte = exceso
        If recorte > Len(cliente) - 5 Then recorte = Len(cliente) - 5
        If recorte > 0 Then cliente = Trim$(Left$(cliente, Len(cliente) - recorte))
    End If
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

