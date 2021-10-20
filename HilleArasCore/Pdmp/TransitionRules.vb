Namespace Pdmp

    Public Class TransitionRules

        Public Function GetTransition(parser As ITransitionParser, targetMaturityState As String, partClassification As String, stage As String, ByRef release As Boolean) As Transition
            Return parser.GetTranstion(targetMaturityState, partClassification, stage, release)
        End Function

    End Class

End Namespace