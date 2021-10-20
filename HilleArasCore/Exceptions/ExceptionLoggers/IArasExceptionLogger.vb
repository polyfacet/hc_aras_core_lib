Imports Aras.IOM

Public Interface IArasExceptionLogger
    Sub Log(inn As Innovator, method As String, message As String, innerException As Exception)

End Interface
