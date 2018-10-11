Public Class XmlOutputFactory

    Public Shared Function GetImplemenation(name As String) As IXmlOutput
        ' TODO: Implement to support different implementations
        Dim xmlOutput As IXmlOutput = Nothing

        Select Case name
            Case "ECO"
                ' TODO: Implement for ECO context
        End Select

        If xmlOutput Is Nothing Then
            ' Set default
            xmlOutput = New BomCompare.OutputFormat.DefaultXmlOutput
        End If

        Return xmlOutput
    End Function

End Class
