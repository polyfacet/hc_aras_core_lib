Imports Aras.IOM
Public Class Utilities

    Private Const MODIFIED_BY_ID As String = "MODIFIED_BY_ID"

    Private ReadOnly Property Inn As Innovator
    Public Sub New(inn As Innovator)
        Me.Inn = inn
    End Sub

    Public Sub SetModifiedBy(item As Item, userId As String)
        Dim table As String = item.getType
        Dim id As String = item.getID

        Dim sb As Text.StringBuilder = New Text.StringBuilder()
        sb.AppendFormat("UPDATE {0} ", table)
        sb.AppendFormat("SET {0} = '{1}'", MODIFIED_BY_ID, userId)
        sb.AppendFormat(" WHERE ID = '{0}'", id)
        Dim sql As String = sb.ToString
        Inn.applySQL(sql)
    End Sub

    Public Sub SetModifiedBy(item As Item, user As Item)
        SetModifiedBy(item, user.getID)
    End Sub

    Public Shared Function GetArasDate(datetime As Date) As String
        Return datetime.ToString("yyyy-MM-ddThh:mm:ss")
    End Function

    ''' <summary>
    ''' Converts Aras date format to Date, returns MinValue if parsing fails
    ''' </summary>
    ''' <param name="arasDate"></param>
    ''' <returns></returns>
    Public Shared Function GetDateFromArasDate(arasDate As String) As Date
        Dim dateResult As Date
        If Not DateTime.TryParseExact(arasDate, "yyyy-MM-ddTHH:mm:ss", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, dateResult) Then
            dateResult = DateTime.MinValue
        End If
        Return dateResult
    End Function


End Class
