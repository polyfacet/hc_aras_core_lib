
Namespace Log


    Public Class TempFileLogger
        Implements ILog


        Private logDirField As String = String.Empty
        Private Property LogDir As String
            Get
                If String.IsNullOrEmpty(logDirField) Then
                    logDirField = GetLogDir
                End If
                Return logDirField
            End Get
            Set(value As String)
                logDirField = value
            End Set
        End Property

        Private filePathField As String = String.Empty
        Public ReadOnly Property FilePath As String
            Get
                If filePathField = String.Empty Then
                    filePathField = IO.Path.Combine(LogDir, Date.Now.Ticks.ToString & ".log")
                End If
                Return filePathField
            End Get
        End Property


        Public Sub LogMessage(message As String) Implements ILog.LogMessage
            Using sw As New IO.StreamWriter(FilePath, True)
                sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
                sw.WriteLine(message)
            End Using
        End Sub

        Private Function GetLogDir() As String
            Dim dir As String = IO.Path.GetTempPath & "\" & System.Reflection.Assembly.GetExecutingAssembly.GetName.FullName
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If
            Return dir
        End Function

    End Class

End Namespace