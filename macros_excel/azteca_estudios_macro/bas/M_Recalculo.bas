Attribute VB_Name = "M_Recalculo"
' Autores: Diana Berumen Estrada
' Fernando Dorantes Nieto
'
' MÓDULO: M_Recalculo
' Propósito: recalcular automáticamente los subtotales de bloque y los totales
' generales cuando se editan o vacían celdas de productos/fechas en la hoja
' Cotización.
'
' Depende de:
'   HojaCotizacion(), FILA_INICIO (M_UUID)
'   suma (Módulo9)

Option Explicit

Public Sub RecalcularTrasCambio(ByVal Target As Range)
    Dim ws As Worksheet
    Dim filaFin As Long
    Dim rango As Range
    Dim eventos As Boolean
    Dim pantalla As Boolean
    Static ocupado As Boolean

    If ocupado Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub

    filaFin = BuscarFilaSubtotal(ws)
    If filaFin <= FILA_INICIO Then Exit Sub

    Set rango = ws.Range("A" & FILA_INICIO & ":M" & (filaFin - 1))
    Set rango = Union(rango, ws.Range("O" & FILA_INICIO & ":O" & (filaFin - 1)))

    If Intersect(Target, rango) Is Nothing Then Exit Sub

    ocupado = True
    eventos = Application.EnableEvents
    pantalla = Application.ScreenUpdating
    On Error GoTo salir

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculate
    Call suma

salir:
    Application.EnableEvents = eventos
    Application.ScreenUpdating = pantalla
    ocupado = False
End Sub

' Macro manual opcional para recalcular todo desde cualquier hoja.
Public Sub RecalcularTodo()
    Dim ws As Worksheet
    Dim eventos As Boolean
    Dim pantalla As Boolean

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    ws.Activate
    eventos = Application.EnableEvents
    pantalla = Application.ScreenUpdating
    On Error GoTo salir

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculate
    Call suma

salir:
    Application.EnableEvents = eventos
    Application.ScreenUpdating = pantalla
End Sub

Private Function BuscarFilaSubtotal(ByVal ws As Worksheet) As Long
    Dim ultima As Long, r As Long

    ultima = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "I").End(xlUp).Row > ultima Then
        ultima = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    End If

    ' 2026-11-09: rótulo por prefijo exacto (EtiquetaTotal, M_Blindaje).
    For r = FILA_INICIO To ultima + 20
        If EtiquetaTotal(ws.Cells(r, "A").Value) = "SUBTOTAL" Or _
           EtiquetaTotal(ws.Cells(r, "I").Value) = "SUBTOTAL" Then
            BuscarFilaSubtotal = r
            Exit Function
        End If
    Next r
End Function
