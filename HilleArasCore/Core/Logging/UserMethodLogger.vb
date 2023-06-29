Imports Aras.IOM

Namespace Log


    Public Class UserMethodLogger

        Private Inn As Innovator
        Private SB As Text.StringBuilder
        Public ReadOnly Property InstanceId As String
        Public ReadOnly Property UserId As String
        Public ReadOnly Property BufferMode As Boolean


        Private Shared Loggers As Dictionary(Of String, UserMethodLogger)
        Private Shared _instance As UserMethodLogger


        Private Sub New(inn As Innovator, uid As String)
            Me.Inn = inn
            Me.InstanceId = uid
            Me.UserId = inn.getUserID
            If UserId <> uid Then
                BufferMode = True
            End If
            Me.SB = New Text.StringBuilder
        End Sub


        Public Shared Function GetLogger(inn As Innovator) As UserMethodLogger
            Return GetLogger(inn, inn.getUserID)
        End Function

        Public Shared Function GetLogger(inn As Innovator, uid As String) As UserMethodLogger
            If Loggers Is Nothing Then
                Loggers = New Dictionary(Of String, UserMethodLogger)
            End If

            If Loggers.ContainsKey(uid) Then
                Console.WriteLine("Using existing instance")
                _instance = Loggers.Item(uid)
            Else
                Console.WriteLine("Creating new instance")
                _instance = New UserMethodLogger(inn, uid)
                Loggers.Add(uid, _instance)
            End If
            Console.WriteLine("Instance Id: " & _instance.InstanceId)
            Return _instance
        End Function

        Public Sub Append(message As String)
            Append(message, -1)
        End Sub

        Public Sub Append(message As String, elapsedTime As Double)
            SB.Append(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            If BufferMode Then
                SB.Append(" (B) ")
            End If
            SB.Append(" : ")
            SB.AppendLine(message)

            If elapsedTime <> -1 Then
                SB.Append(vbTab)
                SB.AppendLine("Execution Time: " & elapsedTime)
            End If
            If Not BufferMode Then
                WriteBufferToFile()
            End If
        End Sub

        Public Sub WriteMessageToFile(message As String)
            Append(message)
            WriteBufferToFile()
        End Sub

        Public Sub WriteBufferToFile()
            Dim fileName As String
            If InstanceId <> UserId Then
                fileName = UserId & "_" & InstanceId & ".log"
            Else
                fileName = UserId & ".log"
            End If
            Dim filePath As String = IO.Path.Combine(LogDir, fileName)
            Using sw As New IO.StreamWriter(filePath, True)
                sw.Write(SB)
            End Using
            Me.SB = New Text.StringBuilder
        End Sub

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

        Private Function GetLogDir() As String
            Dim dir As String = IO.Path.GetTempPath & "\" & System.Reflection.Assembly.GetExecutingAssembly.GetName.FullName
            If Not IO.Directory.Exists(dir) Then
                IO.Directory.CreateDirectory(dir)
            End If
            Return dir
        End Function


    End Class

End Namespace