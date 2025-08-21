'Esta función se podrá usar en las celdas para extraer el número o números fácilmente de un texto
Function GetNumeric(CellRef As String)
    Dim StringLength As Integer
    StringLength = Len(CellRef)
    For i = 1 To StringLength
    If IsNumeric(Mid(CellRef, i, 1)) Then Result = Result & Mid(CellRef, i, 1)
    Next i
    GetNumeric = Result
End Function


'Este código es para ordenas las hojas de un libro alfabéticamente
Sub SortSheetsTabName()
    Application.ScreenUpdating = False
    Dim ShCount As Integer, i As Integer, j As Integer
    ShCount = Sheets.Count
    For i = 1 To ShCount - 1
    For j = i + 1 To ShCount
    If Sheets(j).Name < Sheets(i).Name Then
    Sheets(j).Move before:=Sheets(i)
    End If
    Next j
    Next i
    Application.ScreenUpdating = True
End Sub