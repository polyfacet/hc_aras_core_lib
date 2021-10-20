Namespace BomCompare
    Public Class RelationshipProperty
        Inherits BaseProperty

        Public Sub New(ByVal propertyName As String, ByVal label As String)
            Me.PropertyName = propertyName
            Me.PropertyLabel = label
        End Sub

        Public Overrides ReadOnly Property IsRelationProperty As Boolean
            Get
                Return True
            End Get
        End Property

        Public Overrides Property PropertyLabel As String

        Public Overrides ReadOnly Property PropertyName As String

    End Class
End Namespace