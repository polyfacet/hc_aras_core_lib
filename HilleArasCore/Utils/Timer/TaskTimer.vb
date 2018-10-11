Namespace Utils.Timer


    Public Class TaskTimer

        Private lastBaseLineTick As Long
        Private lastStopTimeTick As Long
        Private initTimeTick As Long

        Private lastTaskResult As TaskResult
        Private Id As String = String.Empty

        ''' <summary>
        ''' Starts a new task timer
        ''' </summary>
        Public Sub New(id As String)
            Me.New()
            Me.Id = id
            lastTaskResult = New TaskResult(id, String.Empty, lastBaseLineTick)
        End Sub

        ''' <summary>
        ''' Starts a new task timer
        ''' </summary>
        Public Sub New()
            lastBaseLineTick = Date.Now.Ticks
            initTimeTick = lastBaseLineTick
        End Sub


        Public Function Start(description As String) As Double

            If lastTaskResult Is Nothing Then
                Dim newTick As Long = Date.Now.Ticks
                lastBaseLineTick = newTick
                lastTaskResult.SetEndTick(newTick)
                Dim t As Double = lastTaskResult.GetMs
                lastTaskResult = New TaskResult(Id, description, newTick)
                Return t
            Else
                Return Start()
            End If

        End Function

        ''' <summary>
        ''' Returns time (ms) passed in ms since last Start. Start/Restart
        ''' </summary>
        ''' <returns></returns>
        Public Function Start() As Double
            Dim newTick As Long = Date.Now.Ticks
            Dim diff As Long = newTick - lastBaseLineTick
            Dim timeElapsedMs As Double = GetMs(newTick, lastBaseLineTick)
            lastBaseLineTick = newTick
            Return timeElapsedMs
        End Function


        ''' <summary>
        ''' Returns time (ms) passed in ms since last Start. Start/Restart
        ''' </summary>
        ''' <returns></returns>
        Public Function Restart(description As String) As Double
            Return Start(description)
        End Function

        ''' <summary>
        ''' Returns time (ms) passed in ms since last Start. Start/Restart
        ''' </summary>
        ''' <returns></returns>
        Public Function Restart() As Double
            Return Start()
        End Function

        ''' <summary>
        ''' Returns time passed in ms since last Start
        ''' </summary>
        ''' <returns></returns>
        Public Function GetElapsedTimeInMs() As Double
            lastStopTimeTick = Date.Now.Ticks
            Return GetMs(lastStopTimeTick, lastBaseLineTick)
        End Function

        Public Function GetTotalTimeInMs() As Double
            Return GetMs(Date.Now.Ticks, initTimeTick)
        End Function

        ''' <summary>
        ''' Returns time passed in ms since last Start
        ''' </summary>
        ''' <returns></returns>
        Public Function [Stop]() As Double
            Return GetElapsedTimeInMs()
        End Function


        Private Function GetMs(endTime As Long, startTime As Long) As Double
            Dim timeSpan As New TimeSpan(endTime - startTime)
            Return timeSpan.TotalMilliseconds
        End Function

    End Class

End Namespace