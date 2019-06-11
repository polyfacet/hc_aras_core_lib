Imports Aras.IOM

Public Class ArasMethodWrapper

    Public Delegate Function ArasMethod(itemMehod As Item) As Item
    Public Shared Function RunArasMethod(itemMethod As Item, arasMethod As ArasMethod)
        Dim callingClass As String = New System.Diagnostics.StackTrace().GetFrame(1).GetMethod.DeclaringType.FullName
        Dim inn As Innovator = itemMethod.getInnovator()
        Dim log As Log.UserMethodLogger = Core.Log.UserMethodLogger.GetLogger(inn)
        log.Append("Run aras method: " & callingClass)
        Dim t As New Utils.Timer.TaskTimer()
        Try
            Dim result As Item = arasMethod.DynamicInvoke(itemMethod)
            Return result
        Catch ex As Exception
            Dim innerEx As Exception = ex.InnerException
            Dim exceptionLogger As New DefaultArasExceptionLogger(inn, "Error")
            exceptionLogger.Log(inn, callingClass, innerEx.Message, innerEx)
            log.Append(innerEx.ToString)
            Return inn.newError(innerEx.Message)
        Finally
            log.Append("Execution time (ms): " & t.GetElapsedTimeInMs())
        End Try

    End Function

End Class
