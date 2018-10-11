Imports Aras.IOM

Public Class ArasException
    Inherits Exception

    Private loggerField As IArasExceptionLogger
    Protected Property Logger As IArasExceptionLogger
        Get
            If loggerField Is Nothing Then
                loggerField = New DefaultArasExceptionLogger(Inn, Me.Classification)
            End If
            Return loggerField
        End Get
        Set(value As IArasExceptionLogger)
            loggerField = value
        End Set
    End Property

    Protected Property Inn As Innovator

    Protected Overridable Property Classification As String

    Protected Sub New(inn As Innovator, ByVal methodName As String, message As String, innerEx As Exception, logger As IArasExceptionLogger)
        MyBase.New(message, innerEx)
        Me.Inn = inn
        Me.Classification = "Error"
        If Not logger Is Nothing Then
            Me.Logger = logger
        End If
        LogError(inn, methodName, message, innerEx)
    End Sub

    Public Sub New(inn As Innovator, ByVal methodName As String, message As String, innerEx As Exception)
        Me.New(inn, methodName, message, innerEx, Nothing)
    End Sub


    Public Sub New(inn As Innovator, ByVal messsage As String)
        Me.New(inn, String.Empty, messsage, Nothing)
    End Sub

    Public Sub New(inn As Innovator, ByVal method As String, ByVal messsage As String)
        Me.New(inn, method, messsage, Nothing)
    End Sub

    Public Sub New(inn As Innovator, message As String, ByVal innerEx As Exception)
        Me.New(inn, String.Empty, message, innerEx)
    End Sub


    Private Sub LogError(inn As Innovator, methodName As String, message As String, innerEx As Exception)
        Logger.Log(inn, methodName, message, innerEx)
    End Sub



End Class
