Public Class XmlOutputFactory

    Public Shared Function GetImplemenation(name As String) As IXmlOutput
        ' TODO: This could be improved, without the hardcoding of "ECO"
        Dim xmlOutput As IXmlOutput = Nothing

        Select Case name
            Case "ECO"
                xmlOutput = New EcoXmlOutput
        End Select

        If xmlOutput Is Nothing Then
            ' Set default
            xmlOutput = New BomCompare.OutputFormat.DefaultXmlOutput
        End If

        Return xmlOutput
    End Function

End Class
