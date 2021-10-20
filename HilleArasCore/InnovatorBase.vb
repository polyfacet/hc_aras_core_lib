Imports Aras.IOM

Public MustInherit Class InnovatorBase

    Public Sub New(ByVal innovator As Innovator)
        Me.Innovator = innovator
    End Sub

    Protected ReadOnly Property Innovator As Innovator

End Class
