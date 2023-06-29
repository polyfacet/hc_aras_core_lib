Imports Aras.IOM

Namespace Log
    Public Class ArasTempFileLogger
        Implements ILog

        Private Inn As Innovator

        Public Sub New(inn As Innovator)
            Me.Inn = inn
        End Sub


        Public Sub LogMessage(message As String) Implements ILog.LogMessage
            Dim fileName As String = Inn.getUserID & ".log"
            Dim filePath As String = IO.Path.Combine(LogDir, fileName)
            Using sw As New IO.StreamWriter(filePath, True)
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                sw.WriteLine(message)
            End Using
        End Sub

        Private logDirField As String = String.Empty
        Private Property LogDir As String
            Get
                If String.IsNullOrEmpty(logDirField) Then
                    logDirField = GetLogDir()
                End If
                Return logDirField
            End Get
            Set(value As String)
                logDirField = value
            End Set
        End Property

        Private Function GetLogDir() As String
            Dim dir As String = IO.Path.GetTempPath & "\" & System.Reflection.Assembly.GetExecutingAssembly.GetName.FullName
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If
            Return dir
        End Function


    End Class
End Namespace
