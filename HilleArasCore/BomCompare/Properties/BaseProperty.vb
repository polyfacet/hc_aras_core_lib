Namespace BomCompare
    Public MustInherit Class BaseProperty
        Implements IBomCompareItemProperty


        MustOverride ReadOnly Property IsRelationProperty As Boolean Implements IBomCompareItemProperty.IsRelationProperty
        MustOverride ReadOnly Property PropertyName As String Implements IBomCompareItemProperty.PropertyName
        MustOverride Property PropertyLabel As String Implements IBomCompareItemProperty.PropertyLabel

        Private isOverrideableField As Boolean = True
        Overridable Property IsComparable As Boolean Implements IBomCompareItemProperty.IsComparable
            Get
                Return isOverrideableField
            End Get
            Set(value As Boolean)
                isOverrideableField = value
            End Set
        End Property

        Public Property Value As String Implements IBomCompareItemProperty.Value

        Private isVisibleField As Boolean = True
        Overridable Property Visible As Boolean Implements IBomCompareItemProperty.Visible
            Get
                Return isVisibleField
            End Get
            Set(value As Boolean)
                isVisibleField = value
            End Set
        End Property

        Public Property Changed As Boolean Implements IBomCompareItemProperty.Changed

        Public Overridable Overloads Function Equals(obj As Object) As Boolean
            Dim otherProperty As BaseProperty = TryCast(obj, BaseProperty)
            If Not otherProperty Is Nothing Then
                If otherProperty.PropertyName = Me.PropertyName Then
                    Return True
                End If
            End If
            Return False
        End Function

    End Class
End Namespace