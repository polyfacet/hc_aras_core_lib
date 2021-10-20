
Namespace Pdmp

    Public Class Transition

        Public Sub New(fromState As String, toState As String)
            Me.FromState = fromState
            Me.ToState = toState
        End Sub



        Public Property FromState As String
        Public Property ToState As String

    End Class

End Namespace
