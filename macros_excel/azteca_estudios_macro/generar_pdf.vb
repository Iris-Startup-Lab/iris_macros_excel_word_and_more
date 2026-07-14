' Autores: Azeneth Guadalupe Garcia Mendez
' Javier Flores Martinez
Sub GenerarPDF()

    Dim ruta As String

    ruta = ThisWorkbook.Path & "\Cotizacion.pdf"

    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
        Filename:=ruta, _
        Quality:=xlQualityStandard, _
        OpenAfterPublish:=True

End Sub

