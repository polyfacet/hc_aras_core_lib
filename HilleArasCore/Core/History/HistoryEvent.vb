Imports Aras.IOM

Public Class HistoryEvent

    Public ReadOnly Property [When] As Date
    Public ReadOnly Property CreatedOn As String
    Public ReadOnly Property WhoName As String
    Public ReadOnly Property WhoId As String
    Public ReadOnly Property Action As String
    Public ReadOnly Property Comment As String
    Public ReadOnly Property ItemState As String
    Public ReadOnly Property Revision As String
    Public ReadOnly Property Generation As String
    Public ReadOnly Property Tick As Long

    Public Sub New(historyItem As Item)
        Me.When = CDate(historyItem.getProperty("created_on"))
        Me.CreatedOn = historyItem.getProperty("created_on")
        Me.WhoName = historyItem.getPropertyAttribute("created_by_id", "keyed_name", "N/A")
        Me.WhoId = historyItem.getProperty("created_by_id", "N/A")
        Me.Action = historyItem.getProperty("action", "N/A")
        Me.Comment = historyItem.getProperty("comments", "")
        Me.ItemState = historyItem.getProperty("item_state", "N/A")
        Me.Revision = historyItem.getProperty("item_major_rev", "N/A")
        Me.Generation = historyItem.getProperty("item_version", "N/A")
        Me.Tick = CLng(historyItem.getProperty("created_on_tick", "0"))
    End Sub

End Class