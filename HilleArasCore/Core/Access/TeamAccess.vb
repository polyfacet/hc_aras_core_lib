Imports Aras.IOM

Namespace Access

    Public Class TeamAccess

        Public Sub New(item As Item)
            DomainItem = item
            Inn = DomainItem.getInnovator
        End Sub

        Private ReadOnly Property Inn As Innovator
        Private Property DomainItem As Item

        Public Function AddAccess(identityName As String, roleName As String) As Item
            Dim teamId As String = DomainItem.getProperty("team_id")
            If String.IsNullOrEmpty(teamId) Then Return Inn.newError("No team")
            Dim identity As Item = GetIdentityByName(identityName)
            If identity.isError Then Return Inn.newError($"Could not find '{identityName}'")
            Dim role As Item = GetIdentityByName(roleName)
            If role.isError Then Return Inn.newError($"Could not find '{roleName}'")
            Return AddAccess(teamId, identity, role)
        End Function

        Public Function RemoveAccess(identityName As String) As Item
            Dim teamId As String = DomainItem.getProperty("team_id")
            If String.IsNullOrEmpty(teamId) Then Inn.newError("No team")
            Dim identity As Item = GetIdentityByName(identityName)
            If identity.isError Then Inn.newError($"Could not find '{identityName}'")
            Return RemoveAccess(teamId, identity)
        End Function

        Public Overrides Function ToString() As String
            If DomainItem Is Nothing Or DomainItem.isError Or DomainItem.isCollection Then Return MyBase.ToString
            Dim teamId As String = DomainItem.getProperty("team_id")
            If String.IsNullOrEmpty(teamId) Then
                Dim amlItem As String = $"<AML><Item action='get' type='{DomainItem.getType}' id='{DomainItem.getID()} /></AML>"
                DomainItem = Inn.applyAML(amlItem)
            End If
            If String.IsNullOrEmpty(teamId) Then Return ("No team")
            Dim sb As New Text.StringBuilder
            Dim teamIdentities As Item = GetTeamIdentities(teamId)
            For i As Integer = 0 To teamIdentities.getItemCount - 1
                Dim teamIdentitiy As Item = teamIdentities.getItemByIndex(i)
                Dim identityName As String = teamIdentitiy.getPropertyAttribute("related_id", "keyed_name")
                Dim roleName As String = teamIdentitiy.getPropertyAttribute("team_role", "keyed_name")
                sb.AppendLine(($"Identity: {identityName} Role: {roleName}"))
            Next
            Return sb.ToString
        End Function


        Private Function AddAccess(teamId As String, identity As Item, role As Item) As Item
            Dim aml As String = $"<AML>
            <Item action='add' type='Team Identity'>
                <source_id>{teamId}</source_id>
                <related_id>{identity.getID}</related_id>
                <team_role>{role.getID}</team_role>
            </Item>
        </AML>"
            Return Inn.applyAML(aml)
        End Function

        Private Function RemoveAccess(teamId As String, identity As Item) As Item
            Dim aml As String = $"<AML>
            <Item action='delete' type='Team Identity' where=""[Team_Identity].source_id='{teamId}' AND [Team_Identity].related_id='{identity.getID}'"" />
        </AML>"
            Return Inn.applyAML(aml)
        End Function

        Private Function GetIdentityByName(identityName As String) As Item
            Dim aml As String = $"<AML>
            <Item action='get' type='Identity'>
                <name>{identityName}</name>
            </Item>
        </AML>"
            Return Inn.applyAML(aml)
        End Function

        Private Function GetTeamIdentities(teamId As String) As Item
            Dim amlTeamIdentities As String = $"<AML>
            <Item action='get' type='Team Identity' select='related_id,team_role'>
            <source_id>{teamId}</source_id>
            </Item>
        </AML>"
            Dim teamIdentities As Item = Inn.applyAML(amlTeamIdentities)
            Return teamIdentities
        End Function

    End Class

End Namespace