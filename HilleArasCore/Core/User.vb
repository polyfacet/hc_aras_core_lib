Imports Aras.IOM

Public Class User

    Public Shared Function GetUserIdentity(inn As Innovator) As Item
        Dim aml As String = $"<AML>
            <Item action='get' type='User' id='{inn.getUserID}' select='owned_by_id' /> 
            </AML>"
        Dim user As Item = inn.applyAML(aml)
        If user.isError Then Return user
        Return inn.getItemById("Identity", user.getProperty("owned_by_id", "N/A"))
    End Function

End Class
