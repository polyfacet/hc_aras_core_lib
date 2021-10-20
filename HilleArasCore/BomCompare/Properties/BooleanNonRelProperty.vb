Namespace BomCompare
    Public Class BooleanNonRelProperty
        Inherits NonRelProperty
        Implements IBooleanBomCompareItemProperty

        Public Sub New(propertyName As String, label As String)
            MyBase.New(propertyName, label)
        End Sub

        Public Function TranslateBooleanStringValue(arasValue As String) As String Implements IBooleanBomCompareItemProperty.TranslateBooleanStringValue
            Return BooleanUtils.ConvertOneZeroToYesNo(arasValue)
        End Function
    End Class
End Namespace
