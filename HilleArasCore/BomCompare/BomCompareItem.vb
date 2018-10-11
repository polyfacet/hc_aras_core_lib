

Namespace BomCompare
    Public Class BomCompareItem

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