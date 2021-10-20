
Namespace Pdmp
    Public Class PdmpQueries
        Implements IDisposable

        Private Connection As SqlClient.SqlConnection

        Private Const CHECKED_BY_VARIABLE_ID As Integer = 49

        Private _connectionString As String
        Public Sub New(connectionString As String)
            _connectionString = connectionString
            Connection = New SqlClient.SqlConnection(_connectionString)
        End Sub

        Private Function GetDataTable(selectCmd As String) As DataTable
            Dim cmd As New SqlClient.SqlCommand(selectCmd)
            cmd.Connection = Connection
            cmd.CommandType = CommandType.Text
            Dim dataAdapter As New SqlClient.SqlDataAdapter(cmd)
            Dim dataTable As New DataTable

            If Connection.State <> ConnectionState.Open Then
                Connection.Open()
            End If

            dataAdapter.Fill(dataTable)

            Return dataTable
        End Function

        Public Function GetDocumentIdByFileName(fileName As String) As Integer
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT DocumentID, Filename FROM Documents ")
            sb.Append(" WHERE Filename = '{0}'")
            cmd = String.Format(sb.ToString, fileName)

            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CInt(dt.Rows(0).Item(0))
            End If
            Return -1
        End Function

        Public Function GetCurrentStatusId(fileId As Integer) As Integer
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT CurrentStatusId, Filename FROM Documents ")
            sb.Append(" WHERE DocumentID = {0}")
            cmd = String.Format(sb.ToString, fileId)

            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CInt(dt.Rows(0).Item(0))
            End If
            Return -1
        End Function

        Public Function GetCurrentStatus(fileId As Integer) As String
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT Name, CurrentStatusId, Filename FROM Documents ")
            sb.Append(" Left Join Status ON Documents.CurrentStatusID = Status.StatusID")
            sb.Append(" WHERE DocumentID = {0}")
            cmd = String.Format(sb.ToString, fileId)
            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CType(dt.Rows(0).Item(0), String)
            End If
            Return ""
        End Function

        Public Function GetFoderId(docId As Integer) As Integer
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT ProjectID FROM DocumentsInProjects ")
            sb.Append(" WHERE DocumentID = '{0}'")
            cmd = String.Format(sb.ToString, docId)

            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CInt(dt.Rows(0).Item(0))
            End If
            Return -1
        End Function

        Public Function GetDocumentLastRevisionNo(fileName As String) As Integer
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("SELECT LatestRevisionNo FROM Documents ")
            sb.Append(" WHERE Filename = '{0}'")
            cmd = String.Format(sb.ToString, fileName)

            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CInt(dt.Rows(0).Item(0))
            End If
            Return -1
        End Function

        Public Function GetDocumentCheckedBySign(docId As Integer) As String
            Dim cmd As String
            Dim sb As New Text.StringBuilder()
            sb.Append("Select ValueText FROM VariableValue ")
            sb.Append("Left Join Documents on VariableValue.DocumentID = Documents.DocumentID ")
            sb.Append("Where VariableID = {0} And Documents.DocumentID = {1} Order By RevisionNo DESC")
            cmd = String.Format(sb.ToString, CHECKED_BY_VARIABLE_ID, docId)
            Dim dt As DataTable = GetDataTable(cmd)
            If dt.Rows().Count > 0 Then
                Return CType(dt.Rows(0).Item(0), String)
            End If
            Return ""
        End Function




#Region "IDisposable Support"
        Private disposedValue As Boolean ' To detect redundant calls

        ' IDisposable
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' TODO: dispose managed state (managed objects).
                End If

                ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
                ' TODO: set large fields to null.
                If Not Connection Is Nothing AndAlso Connection.State <> ConnectionState.Closed Then
                    Connection.Close()
                End If
            End If
            disposedValue = True
        End Sub

        ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
        'Protected Overrides Sub Finalize()
        '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        '    Dispose(False)
        '    MyBase.Finalize()
        'End Sub

        ' This code added by Visual Basic to correctly implement the disposable pattern.
        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
            Dispose(True)
            ' TODO: uncomment the following line if Finalize() is overridden above.
            ' GC.SuppressFinalize(Me)
        End Sub
#End Region



    End Class
End Namespace