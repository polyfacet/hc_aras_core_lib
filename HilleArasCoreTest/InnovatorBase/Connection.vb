
Imports Aras.IOM

Public Class Connection


    Friend url As String = "http://localhost/InnovatorServer"
    Private db As String = "InnovatorSolutions"
    Friend arasUser As String = "admin"
    Friend password As String = "innovator"

    Private conn As HttpServerConnection
    Private inn As Innovator

    Public Sub New(ByVal url As String, ByVal db As String, ByVal userLogin As String, ByVal password As String)
        Me.url = url
        Me.db = db
        Me.arasUser = userLogin
        Me.password = password
    End Sub

    Public Sub New()

    End Sub


    Public ReadOnly Property GetInnovator() As Innovator
        Get
            If inn Is Nothing Then
                inn = IomFactory.CreateInnovator(GetArasConnection)
            End If
            Return inn
        End Get
    End Property

    Public ReadOnly Property GetArasConnection As HttpServerConnection
        Get
            If conn Is Nothing Then
                conn = IomFactory.CreateHttpServerConnection(url, db, arasUser, password)
                Dim login_result As Item = conn.Login()
                If login_result.isError Then
                    conn = Nothing
                    Throw New Exception("Login failed:" & login_result.ToString)
                End If
            End If
            Return conn
        End Get
    End Property

    Public Function Connect() As Boolean
        If GetInnovator IsNot Nothing Then
            Return True
        End If
        Return False
    End Function

    Public Sub Disconnect()
        If GetInnovator IsNot Nothing Then
            conn.Logout()
            inn = Nothing
            conn = Nothing
        End If
    End Sub


End Class
