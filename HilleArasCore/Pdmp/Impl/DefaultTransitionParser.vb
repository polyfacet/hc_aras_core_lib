Imports System.Xml

Namespace Pdmp

    Public Class DefaultTransitionParser
        Implements ITransitionParser



        Private xmlDocField As XmlDocument
        Private Property XmlDoc As XmlDocument
            Get
                If xmlDocField Is Nothing Then
                    xmlDocField = New XmlDocument
                End If
                Return xmlDocField
            End Get
            Set(value As XmlDocument)
                xmlDocField = value
            End Set
        End Property


        Public Sub SetXmlString(xml As String) Implements ITransitionParser.SetXmlString
            XmlDoc.LoadXml(xml)
        End Sub

        Public Function GetTranstion(targetMaturityState As String, partClassification As String, stage As String, release As Boolean) As Transition Implements ITransitionParser.GetTranstion
            Dim maturityElement As XmlElement = GetMatchingRule(targetMaturityState, partClassification, release)
            If maturityElement Is Nothing Then
                Throw New MyArasException(Nothing, "No pdmp rules found matching maturity: " & targetMaturityState & " and part classificaion: " & partClassification)
            End If

            Dim transition As Transition = GetTranstion(maturityElement, stage)

            If transition Is Nothing Then
                Throw New MyArasException(Nothing, "No pdmp rules found for stage: " & stage & " in maturity: " & targetMaturityState & " for classification: " & partClassification)
            End If

            Return transition
        End Function

        Private Function GetMatchingRule(targetMaturityState As String, partClassification As String, ByVal release As Boolean) As XmlElement
            Dim maturityNodes As XmlNodeList
            If release Then
                maturityNodes = Me.XmlDoc.GetElementsByTagName("Released")
            Else
                maturityNodes = Me.XmlDoc.GetElementsByTagName("ToMaturityState")
            End If

            For Each maturityNode As XmlNode In maturityNodes
                If TypeOf maturityNode Is XmlElement Then
                    Dim maturityElement As XmlElement = CType(maturityNode, XmlElement)
                    ' Check if it is the specified target maturirty
                    If maturityElement.HasAttribute("value") _
                        AndAlso maturityElement.GetAttribute("value") = targetMaturityState Then
                        ' OK check if it is the specfied part classification
                        Dim classificationNodes As XmlNodeList = maturityElement.SelectNodes("PartClassifications/Classification")
                        For Each classificationNode As XmlNode In classificationNodes
                            If TypeOf classificationNode Is XmlElement Then
                                Dim classifactionElement As XmlElement = CType(classificationNode, XmlElement)
                                ' Check if it is the input part classification
                                If classifactionElement.InnerText = partClassification Then
                                    ' Found the matching maturity element!
                                    Return maturityElement
                                End If
                            End If
                        Next
                    End If
                End If
            Next

            Return Nothing
        End Function

        Private Function GetTranstion(maturityElement As XmlElement, stage As String) As Transition
            For Each childNode As XmlNode In maturityElement.ChildNodes
                If TypeOf childNode Is XmlElement _
                    AndAlso childNode.Name = "Stages" Then
                    ' Loop the childnodes to match the stage with the xmlelement name
                    For Each stageNode As XmlNode In childNode.ChildNodes
                        If TypeOf stageNode Is XmlElement _
                            AndAlso stageNode.Name = stage Then
                            ' Found the matching xmlelement, get from and to values
                            Dim stageElement As XmlElement = CType(stageNode, XmlElement)
                            Dim errorMessage As String = "Missing {0} attribute in: " & maturityElement.Name & " / " & stageElement.Name
                            Dim fromState As String
                            Dim toState As String
                            If stageElement.HasAttribute("from") Then
                                fromState = stageElement.GetAttribute("from")
                            Else
                                Throw New MyArasException(Nothing, String.Format(errorMessage, "from"))
                            End If
                            If stageElement.HasAttribute("to") Then
                                toState = stageElement.GetAttribute("to")
                            Else
                                Throw New MyArasException(Nothing, String.Format(errorMessage, "to"))
                            End If
                            ' All ok
                            Dim transition As New Transition(fromState, toState)
                            Return transition
                        End If
                    Next
                End If
            Next
            Return Nothing

        End Function
    End Class

End Namespace
