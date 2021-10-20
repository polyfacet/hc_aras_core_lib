Imports System.Runtime.CompilerServices
Imports Aras.IOM

Public Module ItemExt

    Private Const GENERATION As String = "generation"
    Private Const CONFIG_ID As String = "config_id"

    <Extension()>
    Public Function GetLatestReleased(item As Item) As Item
        Dim inn As Innovator = item.getInnovator
        Try
            Dim type As String = item.getType
            Dim configId As String = GetConfigId(item)

            Dim amlQuery As String = $"<AML>
             <Item action='get' type='{type}' serverEvents='0' maxRecords='1' orderBy='generation DESC'>
             <config_id>{configId}</config_id>
             <generation condition='ge'>0</generation>
             <is_released>1</is_released>
             </Item></AML>"
            Dim latestReleased As Item = inn.applyAML(amlQuery)
            Return latestReleased
        Catch ex As Exception
            Return inn.newError(ex.ToString)
        End Try
    End Function

    <Extension()>
    Public Function PreviousReleased(item As Item) As Item
        Dim inn As Innovator = item.getInnovator
        Try
            Dim type As String = item.getType
            Dim configId As String = GetConfigId(item)
            Dim inputGeneration As Integer = GetGeneration(item)

            Dim amlQuery As String = $"<AML>
             <Item action='get' type='{type}' serverEvents='0' orderBy='generation DESC'>
             <config_id>{configId}</config_id>
             <generation condition='ge'>0</generation>
             <is_released>1</is_released>
             </Item></AML>"
            Dim releasedVersions As Item = inn.applyAML(amlQuery)
            For i As Integer = 0 To releasedVersions.getItemCount() - 1
                Dim releasedVersion As Item = releasedVersions.getItemByIndex(i)
                If CInt(releasedVersion.getProperty(GENERATION)) < inputGeneration Then
                    Return releasedVersion
                End If
            Next
            Return inn.newError("No Previous Released found.")
        Catch ex As Exception
            Return inn.newError(ex.ToString)
        End Try
    End Function

    <Extension()>
    Public Function TeamAccess(item As Item) As Access.TeamAccess
        Return New Access.TeamAccess(item)
    End Function

    Private Function GetConfigId(item As Item) As String
        Dim inn As Innovator = item.getInnovator
        Dim type As String = item.getType
        Dim configId As String = item.getProperty(CONFIG_ID)
        If String.IsNullOrEmpty(configId) Then
            configId = FetchProperty(item, CONFIG_ID)
        End If
        Return configId
    End Function

    Private Function GetGeneration(item As Item) As Integer
        Dim inn As Innovator = item.getInnovator
        Dim type As String = item.getType
        Dim gen As String = item.getProperty(GENERATION)
        If String.IsNullOrEmpty(gen) Then
            gen = FetchProperty(item, GENERATION)
        End If
        Return CInt(gen)
    End Function

    Private Function FetchProperty(item As Item, propertyName As String) As String
        Dim inn As Innovator = item.getInnovator
        Dim type As String = item.getType
        Dim value As String = item.getProperty(propertyName)
        If String.IsNullOrEmpty(value) Then
            Dim amlQuery1 As String = $"<AML>
                    <Item action='get' type='{type}' serverEvents='0' id='{item.getID}' select='{propertyName}'>
                    </Item></AML>"
            item = inn.applyAML(amlQuery1)
            value = item.getProperty(propertyName)
        End If
        Return value
    End Function

End Module


