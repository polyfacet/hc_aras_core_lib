Imports Aras.IOM
Public Class Email

    Private ReadOnly Property Inn As Innovator

    Public Sub New(innovator As Innovator)
        Me.Inn = innovator
    End Sub

    Public Function Send(fromIdName As String, subject As String, body As String) As Boolean
        Dim emailItem As Item = Inn.newItem("EMail Message")
        emailItem.setProperty("subject", subject)
        emailItem.setProperty("body_plain", body)
        Dim fromUserItem As Item = GetFromUser()
        emailItem.setPropertyItem("from_user", fromUserItem)

        Dim indentityItem As Item = GetIdentityItem(fromIdName)
        Return emailItem.email(emailItem, indentityItem)
    End Function

    Private Function GetIdentityItem(fromIdName As String) As Item
        Dim iden As Item = Inn.newItem("Identity", "get")
        iden.setProperty("name", fromIdName)
        iden = iden.apply()
        Return iden
    End Function

    Private Function GetFromUser() As Item
        Dim userId As String = Inn.getUserID()
        Return Inn.getItemById("User", userId)
    End Function
End Class
