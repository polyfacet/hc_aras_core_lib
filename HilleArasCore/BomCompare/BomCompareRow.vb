Imports System.Text

Namespace BomCompare

    Public Class BomCompareRow
        Implements IComparable

        Public Enum ChangeTypeEnum
            Unchanged
            Added
            Removed
            Updated
            Replaced
        End Enum

        Public Sub New(ByVal index As String, ByVal changeType As ChangeTypeEnum, ByVal changeDescription As List(Of String), compareItem As BomCompareItem, baseItem As BomCompareItem)
            Me.Index = index
            Me.ChangeType = changeType
            Me.ChangeDescription = changeDescription
            Me.CompareItem = compareItem
            Me.BaseItem = baseItem
        End Sub


        Public Property Index As String

        Property ChangeType As ChangeTypeEnum
        Property ChangeDescription As List(Of String)
        Property CompareItem As BomCompareItem
        Property BaseItem As BomCompareItem

        Public Overrides Function ToString() As String
            Dim sb As New StringBuilder
            sb.AppendLine("Index: " & Index)
            sb.AppendLine("ChangeType: " & ChangeType.ToString)

            sb.AppendLine("Change description:")
            For Each changeDesc In ChangeDescription
                sb.AppendLine(changeDesc)
            Next

            If CompareItem Is Nothing Then
                sb.AppendLine("No compare item")
            Else
                sb.AppendLine("Compare Item:")
                sb.Append(BomCompareItemToString(CompareItem))
            End If

            If BaseItem Is Nothing Then
                sb.AppendLine("No base item")
            Else
                sb.AppendLine("BaseItem Item:")
                sb.Append(BomCompareItemToString(BaseItem))
            End If
            Return sb.ToString()
        End Function

        Private Function BomCompareItemToString(ByVal item As BomCompareItem) As String
            Dim sb As New StringBuilder
            For Each prop As IBomCompareItemProperty In item.BomCompareItemProperties
                sb.Append(prop.PropertyName)
                sb.Append("=")
                sb.Append(prop.Value)
                sb.Append(" Changed?=" & prop.Changed)
                sb.AppendLine()
            Next
            Return sb.ToString()
        End Function

        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            Dim otherBomCompareRow As BomCompareRow = TryCast(obj, BomCompareRow)
            If otherBomCompareRow Is Nothing Then
                Return 0
            End If

            ' Try to sort by integer if possible
            Dim otherIndex As Integer
            Dim thisIndex As Integer
            If Integer.TryParse(otherBomCompareRow.Index, otherIndex) AndAlso Integer.TryParse(Me.Index, thisIndex) Then
                If thisIndex > otherIndex Then
                    Return 1
                ElseIf thisIndex < otherIndex Then
                    Return -1
                End If
                Return 0
            Else
                If Me.Index > otherBomCompareRow.Index Then
                    Return 1
                ElseIf Me.Index < otherBomCompareRow.Index Then
                    Return -1
                End If
                Return 0
            End If


        End Function
    End Class

End Namespace