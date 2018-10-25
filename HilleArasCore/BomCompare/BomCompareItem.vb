

Namespace BomCompare
    Public Class BomCompareItem

        Public Sub New(arasTypeName As String, id As String)
            Me.TypeName = arasTypeName
            Me.Id = id
        End Sub

        Public ReadOnly Property TypeName As String
        Public ReadOnly Property Id As String

        Private bomCompareItemPropertiesField As List(Of IBomCompareItemProperty)
        Property BomCompareItemProperties As List(Of IBomCompareItemProperty)
            Get
                If bomCompareItemPropertiesField Is Nothing Then
                    bomCompareItemPropertiesField = New List(Of IBomCompareItemProperty)
                End If
                Return bomCompareItemPropertiesField
            End Get
            Set(value As List(Of IBomCompareItemProperty))
                bomCompareItemPropertiesField = value
            End Set
        End Property


    End Class
End Namespace