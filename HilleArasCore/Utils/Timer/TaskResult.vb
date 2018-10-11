Namespace Utils.Timer



    Public Class TaskResult

        Public Sub New(id As String, description As String, startTick As Long)
            Me.Id = id
            Me.Description = description
            Me.StartTick = startTick
        End Sub

        Public Sub New(description As String, startTick As Long)
            Me.New(String.Empty, description, startTick)
        End Sub

        Public Sub New(startTick As Long)
            Me.New(String.Empty, startTick)
        End Sub

        Public Property EndTick As Long = 0
        Public ReadOnly Property StartTick As Long
        Public ReadOnly Property Description As String
        Public ReadOnly Property Id As String

        Public Sub SetEndTick(tick As Long)
            Me.EndTick = tick
        End Sub

        Public Function GetMs() As Double
            If EndTick = 0 Then
                Return -1
            End If
            Dim timeSpan As New TimeSpan(EndTick - StartTick)
            Return timeSpan.TotalMilliseconds
        End Function

    End Class


End Namespace