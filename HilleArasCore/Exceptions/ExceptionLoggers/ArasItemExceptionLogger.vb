Imports Aras.IOM

Public Class ArasItemExceptionLogger
    Inherits DefaultArasExceptionLogger

    Private Item As Item
    Public Sub New(inn As Innovator, ByVal item As Item, ByVal classification As String)
        MyBase.New(inn, classification)
        Me.Item = item
    End Sub

    Protected Overrides Sub AppendProperties()
        Me.AppendProperty("item_type", Item.getType)
        Me.AppendProperty("item_id", Item.getID)
    End Sub



End Class
