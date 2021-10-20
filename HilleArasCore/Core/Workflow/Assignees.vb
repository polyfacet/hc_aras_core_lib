Imports Aras.IOM

Namespace Workflow
    Public Class Assignees

        ''' <summary>
        ''' 
        ''' </summary>
        ''' <param name="sourceItem">example ECO, MCO</param>
        Public Sub New(sourceItem As Item)
            Me.SourceItem = sourceItem
            Inn = sourceItem.getInnovator
        End Sub

        Protected Inn As Innovator
        Private SourceItem As Item

        Private assigneedIdentitiesField As List(Of Item)
        Public ReadOnly Property AssigneedIdentities As List(Of Item)
            Get
                If assigneedIdentitiesField Is Nothing Then
                    assigneedIdentitiesField = LoadCurrentAssigneesToVote()
                End If
                Return assigneedIdentitiesField
            End Get
        End Property


        Private nonClosedActivityAssignmentsField As List(Of Item)
        Public ReadOnly Property NonClosedActivityAssignments As List(Of Item)
            Get
                If nonClosedActivityAssignmentsField Is Nothing Then
                    nonClosedActivityAssignmentsField = New List(Of Item)
                    Me.LoadCurrentAssigneesToVote()
                End If
                Return nonClosedActivityAssignmentsField
            End Get
        End Property



        ''' <summary>
        ''' Returns user identities who have not voted in current activity of workflow process
        ''' </summary>
        ''' <returns></returns>
        Public Function CurrentAssigneesToVote() As List(Of Item)
            Return LoadCurrentAssigneesToVote()
        End Function

        ''' <summary>
        ''' Returns user names who have not voted in current activity of workflow process
        ''' </summary>
        ''' <returns></returns>
        Public Function CurrentAssigneesToVoteNames() As List(Of String)
            Dim listNames As New List(Of String)
            Dim assigneedIdentities As List(Of Item) = CurrentAssigneesToVote()
            For Each identity As Item In assigneedIdentities
                Dim name As String = identity.getProperty("keyed_name")
                If Not listNames.Contains(name) Then
                    listNames.Add(name)
                End If
            Next
            Return listNames
        End Function


        ''' <summary>
        ''' 'Get latest creation date from assignments (non closed)
        ''' </summary>
        ''' <returns></returns>
        Public Function AssignementDateForNonClosedActivity() As String
            Dim dateString As String = String.Empty
            For Each assignement As Item In Me.NonClosedActivityAssignments
                Dim creationDateString As String = assignement.getProperty("created_on")
                If dateString = String.Empty Then
                    dateString = creationDateString
                Else
                    ' Compare and set latest date
                    Dim thisDate As Date = Utilities.GetDateFromArasDate(creationDateString)
                    Dim compareDate As Date = Utilities.GetDateFromArasDate(dateString)
                    If thisDate > compareDate Then
                        dateString = creationDateString
                    End If
                End If
            Next
            Return dateString
        End Function

        Private Function LoadCurrentAssigneesToVote() As List(Of Item)
            Dim assigneedIdentities As New List(Of Item)
            ' Get Workflow Process
            Dim query As New Text.StringBuilder
            query.Append("<AML>")
            query.Append("<Item action ='get' type='Workflow' select='related_id'>")
            query.AppendFormat("<source_id>{0}</source_id>", SourceItem.getID)
            query.Append("<related_id>")
            query.Append("<Item type ='Workflow Process'>")
            query.Append("<state>Active</state>")
            query.Append("</Item></related_id></Item></AML>")
            Dim queryResult As Item = Inn.applyAML(query.ToString)

            ' Get active Activity Assigments for workflowprocess
            If Not queryResult.isError Then
                Dim wfpId As String = queryResult.getItemsByXPath("//Item[@type='Workflow Process']").getItemByIndex(0).getID
                query = New Text.StringBuilder
                query.Append("<AML>")
                query.AppendFormat("<Item type = 'Workflow Process' action='get' id='{0}' >", wfpId)
                query.Append("<Relationships>")
                query.Append("<Item action ='get' type='Workflow Process Activity' select='related_id'>")
                query.Append("<related_id>")
                query.Append("<Item type='Activity' action='get' select='id,state'>")
                query.Append("<state>Active</state>")
                query.Append("<Relationships>")
                query.Append("<Item action='get' type='Activity Assignment'>")
                query.Append("<closed_on condition='is null'></closed_on>")
                query.Append("</Item></Relationships></Item></related_id></Item></Relationships></Item>></AML>")
                Dim queryResultActivities As Item = Inn.applyAML(query.ToString)
                If Not queryResultActivities.isError Then
                    Dim activityAssignments As Item = queryResultActivities.getItemsByXPath("//Item[@type='Activity Assignment']")
                    nonClosedActivityAssignmentsField = New List(Of Item)
                    For i As Integer = 0 To activityAssignments.getItemCount - 1
                        nonClosedActivityAssignmentsField.Add(activityAssignments.getItemByIndex(i))
                        Dim identity As Item = activityAssignments.getItemByIndex(i).getRelatedItem
                        assigneedIdentities.Add(identity)
                    Next
                End If
            End If
            Return assigneedIdentities
        End Function

    End Class
End Namespace
