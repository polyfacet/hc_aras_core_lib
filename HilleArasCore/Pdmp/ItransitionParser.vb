

Namespace Pdmp

    Public Interface ITransitionParser

        Sub SetXmlString(ByVal xml As String)
        Function GetTranstion(ByVal targetMaturityState As String, ByVal partClassification As String, ByVal stage As String, ByVal release As Boolean) As Transition

    End Interface

End Namespace