Sub GenerarPDFNombreInteligente()

    Dim ruta As String
    Dim cliente As String
    Dim proyecto As String
    Dim fechaCot As String
    Dim nombreArchivo As String

    cliente = Range("C3").Value
    proyecto = Range("C4").Value
    fechaCot = Format(Range("K3").Value, "dd-mm-yyyy")

    cliente = Replace(cliente, "/", "-")
    proyecto = Replace(proyecto, "/", "-")

    ruta = ThisWorkbook.Path & "\"

    nombreArchivo = "Cotizacion_" & cliente & "_" & proyecto & "_" & fechaCot & ".pdf"

    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ruta & nombreArchivo, _
        Quality:=xlQualityStandard, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True

End Sub

