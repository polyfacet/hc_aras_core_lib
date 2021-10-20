Public Class BooleanUtils

    Public Shared Function ConvertOneZeroToYesNo(value As String) As String
        Return If((value = "1"), "Yes", "No")
    End Function

End Class
