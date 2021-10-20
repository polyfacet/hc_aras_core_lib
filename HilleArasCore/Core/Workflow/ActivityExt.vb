Imports Aras.IOM

Namespace Workflow
    Public Class ActivityExt
        Public Activity As Item

        Public WorkflowProcess As Item

        Public Workflow As Item

        Public SourceItem As Item

        Public ActivityName As String

        Protected Inn As Innovator

        Public Sub New(ByVal activity As Item)
            MyBase.New
            Me.Inn = activity.getInnovator
            Me.Activity = activity
            Me.Init()
        End Sub

        Public Sub New(ByVal activityExt As ActivityExt)
            MyBase.New
            Me.Activity = activityExt.Activity
            Me.WorkflowProcess = activityExt.WorkflowProcess
            Me.Workflow = activityExt.Workflow
            Me.SourceItem = activityExt.SourceItem
            Me.ActivityName = activityExt.ActivityName
            Me.Inn = activityExt.Inn
        End Sub

        Private _activityAssignment As Item
        Public ReadOnly Property ActivityAssignment() As Item
            Get
                If _activityAssignment Is Nothing Then
                    _activityAssignment = FetchActivityAssignment()
                End If
                Return _activityAssignment
            End Get
        End Property

        Private Sub Init()
            Me.ActivityName = Me.Activity.getProperty("name")
            Dim sb As New Text.StringBuilder
            sb.Append("<AML>")
            sb.Append("<Item type=""Workflow Process"" action=""get"">")
            sb.Append("<Relationships>")
            sb.Append("<Item type=""Workflow Process Activity"" action=""get"">")
            sb.Append("<related_id>")
            sb.Append("<Item type=""Activity"" id=""{0}"" action=""get"" />")
            sb.Append("</related_id>")
            sb.Append("</Item>")
            sb.Append("<Item type=""Workflow Process Activity"" action=""get"">")
            sb.Append("<related_id>")
            sb.Append("<Item type=""Activity"" action=""get"">")
            sb.Append("<Relationships>")
            sb.Append("<Item type=""Workflow Process Path""></Item>")
            sb.Append("<Item type=""Activity Assignment""></Item>")
            sb.Append("</Relationships></Item></related_id></Item></Relationships></Item></AML>")

            Dim aml As String = sb.ToString
            aml = String.Format(aml, Me.Activity.getID)
            Me.WorkflowProcess = Me.Inn.applyAML(aml)

            ' Get the Controlled Item , support versionable source item
            Dim amlWfquery As String = "<AML><Item action='get' type='Workflow' where=""[Workflow].related_id='{0}'"" orderBy='created_on DESC'></Item></AML>"
            amlWfquery = String.Format(amlWfquery, Me.WorkflowProcess.getID)
            Me.Workflow = Me.Inn.applyAML(amlWfquery)
            If (Me.Workflow.getItemCount > 1) Then
                ' Get latest
                Me.Workflow = Me.Workflow.getItemByIndex(0)
            End If

            If Not Workflow.isError Then
                ' Get the source item. PCO/ECO/MCO/PPR
                Me.SourceItem = Me.Inn.newItem(Me.Workflow.getPropertyAttribute("source_type", "name"), "get")
                Me.SourceItem.setID(Me.Workflow.getProperty("source_id"))
                Me.SourceItem = Me.SourceItem.apply
            End If

        End Sub

        Private Function FetchActivityAssignment() As Item
            Dim assignmentId As String = Activity.getProperty("AssignmentId")
            If Not String.IsNullOrEmpty(assignmentId) Then
                Dim amlQueryAssignment As String = $"<AML>
                    <Item action='get' type='Activity Assignment' id='{assignmentId}' >
                    </Item></AML>"
                Return Inn.applyAML(amlQueryAssignment)
            End If
            Return Inn.newError("No AssignmentId in Activity")
        End Function

    End Class
End Namespace