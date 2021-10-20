
Namespace BomCompare.OutputFormat

    Public Class DefaultOutputSettings
        Implements IOutputSettings

        Private cellColorsField As IOutputColors
        Public Property CellColors As IOutputColors Implements IOutputSettings.CellColors
            Get
                If cellColorsField Is Nothing Then
                    cellColorsField = New DefaultOutputColors
                End If
                Return cellColorsField
            End Get
            Set(value As IOutputColors)
                cellColorsField = value
            End Set
        End Property

        Private contentItemTypeField As String = "Part"
        Public Property ContentItemType As String Implements IOutputSettings.ContentItemType
            Get
                Return contentItemTypeField
            End Get
            Set(value As String)
                contentItemTypeField = value
            End Set
        End Property
    End Class

End Namespace