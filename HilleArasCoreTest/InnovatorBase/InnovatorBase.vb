
Imports Aras.IOM

Public MustInherit Class InnovatorBase

    Shared Inn As Innovator

    Public Shared Url As String = My.Settings.ARAS_URL 'http://localhost/innovator/
    Public Shared DB As String = My.Settings.ARAS_DB 'InnovatorSolutions
    Public Shared User As String = My.Settings.ARAS_USER 'admin
    Public Shared Password As String = My.Settings.ARAS_PASSWORD 'innovator
    Public Shared TimeoutSeconds As Integer = 0

    Private Shared Sub createInnovator()
        Dim MyConnection As HttpServerConnection = IomFactory.CreateHttpServerConnection(Url, DB, User, Password)
        If TimeoutSeconds > 0 Then
            MyConnection.Timeout = 1000 * TimeoutSeconds
        End If
        Inn = New Innovator(MyConnection)
        MyConnection.Login()
    End Sub

    Public Shared Function getInnovator() As Innovator
        If Inn Is Nothing Then createInnovator()
        Return Inn
    End Function

    Public Shared Function ToStringAddress()
        Return Url & " DB: " & DB & " User:" & User
    End Function

End Class
