Imports Aras.IOM
Imports Hille.Aras.Core.Log

Public Class History

    Private Const ACTION_1 As String = "FormPrint" ' R11SP9: "Only" custom action allowed in aml
    Private Const ACTION_2 As String = "Update"

    Private Property Item() As Item
    Private ReadOnly Property Inn As Innovator

    Private configIdField As String = String.Empty
    Private ReadOnly Property ConfigId As String
        Get
            If configIdField = String.Empty Then
                configIdField = GetConfigId()
            End If
            Return configIdField
        End Get
    End Property


    Public Sub New(item As Item)
        Me.Item = item
        Inn = item.getInnovator
    End Sub

    Public Sub AddHistory(message As String)
        AddHistory(message, ACTION_2)
    End Sub


    Public Sub AddHistory(message As String, action As String)
        Try
            Dim historyContainerId As String = CreateEntry(message)
            If Not String.IsNullOrEmpty(historyContainerId) Then
                ChangeAction(historyContainerId, action, message)
            End If
        Catch ex As Exception
            Dim log As ILog = New TempFileLogger
            log.LogMessage(ex.ToString)
        End Try
    End Sub

    ''' <summary>
    ''' Get all history of an Item
    ''' </summary>
    Public Function GetFullHistory() As List(Of HistoryEvent)
        Dim historyList As New List(Of HistoryEvent)

        Dim sb As New Text.StringBuilder
        sb.Append("<AML>")
        sb.Append("<Item action ='get' type='History' orderBy='created_on_tick'>")
        sb.Append("<source_id><Item type='History Container'>")
        sb.AppendFormat("<item_config_id>{0}</item_config_id>", ConfigId)
        sb.Append("</Item></source_id></Item></AML>")
        Dim amlQuery As String = sb.ToString

        Dim historyEntries As Item = Inn.applyAML(amlQuery)
        For i As Integer = 0 To historyEntries.getItemCount() - 1
            historyList.Add(New HistoryEvent(historyEntries.getItemByIndex(i)))
        Next
        Return historyList
    End Function

    ''' <summary>
    ''' Returns id of History Contatainer
    ''' </summary>
    ''' <param name="message"></param>
    ''' <returns></returns>
    Private Function CreateEntry(message As String) As String
        Dim sb As New Text.StringBuilder
        sb.Append("<AML>")
        sb.AppendFormat("<Item type='{0}' action='AddHistory' id='{1}'>", Item.getType, Item.getID)
        sb.AppendFormat("<action>{0}</action>", ACTION_1)
        sb.AppendFormat("<form_name>{0}</form_name>", message)
        sb.Append("</Item></AML>")
        Dim aml As String = sb.ToString()
        Dim result As Item = Inn.applyAML(aml)
        If result.isError Then
            Return String.Empty
        End If

        Dim sb2 As New Text.StringBuilder
        sb2.Append("<AML><Item action='get' type='History Container' select='id'>")
        sb2.AppendFormat("<item_config_id>{0}</item_config_id>", ConfigId)
        sb2.Append("</Item></AML>")
        result = Inn.applyAML(sb2.ToString)
        If result.isError Then
            Return String.Empty
        End If
        Return result.getID
    End Function

    Private Function GetConfigId() As String
        Dim item_config_id As String = Item.getProperty("config_id", "")
        If String.IsNullOrEmpty(item_config_id) Then
            ' Retieve item config_id
            Dim amlGetItem As String = String.Format("<AML><Item action='get' type='{0}' id='{1}' select='config_id' /></AML>", Item.getType, Item.getID)
            Dim itemTemp As Item = Inn.applyAML(amlGetItem)
            item_config_id = itemTemp.getProperty("config_id", "")
        End If
        Return item_config_id
    End Function

    Private Sub ChangeAction(historyContainerId As String, action As String, message As String)
        Dim sb As Text.StringBuilder = New Text.StringBuilder()
        sb.Append("UPDATE HISTORY ")
        sb.AppendFormat("SET ACTION = '{0}'", action)
        sb.AppendFormat(" WHERE SOURCE_ID = '{0}' AND ITEM_ID = '{1}' AND COMMENTS = '{2}' AND ACTION = '{3}'", historyContainerId, Item.getID, message, ACTION_1)
        Dim sql As String = sb.ToString
        Inn.applySQL(sql)
    End Sub


End Class
