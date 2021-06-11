Option Explicit

Sub Combinar()

    Dim hojaParticipantes As Worksheet
    Dim nombreParticipante As String
    Dim filaInicial As Long
    Dim objetoPowerPoint As Object
    Dim objetoPresentacion As Object
    Dim objetoDiapositiva As Object
    Dim objetoCuadroTexto As Object
    
    Set hojaParticipantes = Worksheets("hoja 1")
    
    Set objetoPowerPoint = CreateObject("Powerpoint.Application")
    objetoPowerPoint.Visible = True
    Set objetoPresentacion = objetoPowerPoint.presentations.Open(ThisWorkbook.Path & "\ejemploCertificado.pptx")
    objetoPresentacion.SaveAs ThisWorkbook.Path & "\combinados.pptx"
    
    filaInicial = 2
    Do While hojaParticipantes.Cells(filaInicial, 1) <> ""
        nombreParticipante = hojaParticipantes.Cells(filaInicial, 1)
        
        Set objetoDiapositiva = objetoPresentacion.slides(1).Duplicate
        For Each objetoCuadroTexto In objetoDiapositiva.Shapes
            If objetoCuadroTexto.HasTextFrame Then
                If objetoCuadroTexto.TextFrame.hastext Then
                    objetoCuadroTexto.TextFrame.TextRange.Replace "<nombre>", nombreParticipante
                End If
            End If
        Next
        filaInicial = filaInicial + 1
    Loop
    objetoPresentacion.slides(1).Delete
    objetoPresentacion.Save
    objetoPresentacion.Close
    
End Sub
