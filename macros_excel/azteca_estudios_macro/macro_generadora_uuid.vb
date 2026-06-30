' Autor: Fernando Dorantes Nieto
' Revisor: DeepSeek V4 Pro

' --- DECLARACIONES DEL SISTEMA WINDOWS ---
' Nota: PtrSafe es necesario para versiones de Excel de 64 bits
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As GUID) As Long
#End If

' --- FUNCIÓN AUXILIAR PARA DAR FORMATO ---
Private Function FormatearGUID(g As GUID) As String
    Dim strGUID As String
    
    ' Convierte los números a Hexadecimal y les da el formato correcto con guiones
    strGUID = Right("00000000" & Hex(g.Data1), 8) & "-" & _
              Right("0000" & Hex(g.Data2), 4) & "-" & _
              Right("0000" & Hex(g.Data3), 4) & "-" & _
              Right("00" & Hex(g.Data4(0)), 2) & Right("00" & Hex(g.Data4(1)), 2) & "-" & _
              Right("00" & Hex(g.Data4(2)), 2) & Right("00" & Hex(g.Data4(3)), 2) & _
              Right("00" & Hex(g.Data4(4)), 2) & Right("00" & Hex(g.Data4(5)), 2) & _
              Right("00" & Hex(g.Data4(6)), 2) & Right("00" & Hex(g.Data4(7)), 2)
              
    ' Devuelve el UUID en mayúsculas (cambia a LCase si lo quieres en minúsculas)
    FormatearGUID = UCase(strGUID)
End Function

' --- FUNCIÓN PRINCIPAL QUE PUEDES USAR EN EXCEL ---
Public Function GENERAR_UUID() As String
    Dim g As GUID
    
    ' Si el sistema genera el GUID correctamente (retorno 0), lo formatea
    If CoCreateGuid(g) = 0 Then
        GENERAR_UUID = FormatearGUID(g)
    Else
        GENERAR_UUID = "Error al generar"
    End If
End Function

' --- MACRO PARA INSERTAR UUIDs EN VARIAS CELDAS SELECCIONADAS ---
Sub InsertarUUIDsEnSeleccion()
    Dim celda As Range
    
    ' Recorre cada celda seleccionada
    For Each celda In Selection
        ' Evita que se ejecute en celdas combinadas o protegidas
        If Not celda.MergeCells And Not celda.Locked Then
            celda.Value = GENERAR_UUID()
            ' Formato de texto para que Excel no borre los ceros o lo convierta en número
            celda.NumberFormat = "@"
        End If
    Next celda
End Sub