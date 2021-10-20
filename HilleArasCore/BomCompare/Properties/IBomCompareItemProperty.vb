Namespace BomCompare
    Public Interface IBomCompareItemProperty

        ReadOnly Property PropertyName As String
        Property PropertyLabel As String
        ReadOnly Property IsRelationProperty As Boolean
        Property IsComparable As Boolean
        Property Visible As Boolean
        Property Value As String
        Property Changed As Boolean

    End Interface
End Namespace