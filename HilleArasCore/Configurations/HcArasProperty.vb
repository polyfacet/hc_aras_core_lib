Imports Aras.IOM
Imports Aras.Server.Core
Imports Aras.Server

Public Class HCArasProperty
    Inherits InnovatorBase

    Private ReadOnly Property PropertyName As String
    Private CCO As CallContext

    Public Sub New(inn As Innovator, cco As CallContext, propertyName As String)
        MyBase.New(inn)
        Me.CCO = cco
        Me.PropertyName = propertyName
    End Sub

    Public Sub New(inn As Innovator, propertyName As String)
        Me.New(inn, Nothing, propertyName)
    End Sub

    Private _item As Item
    Public ReadOnly Property Item As Item
        Get
            If _item Is Nothing Then
                _item = FetchItem()
            End If
            Return _item
        End Get
    End Property

    Public ReadOnly Property Exists As Boolean
        Get
            If Item.isError() Then Return False
            Return True
        End Get
    End Property

    Public ReadOnly Property Enabled As Boolean
        Get
            If CCO Is Nothing Then Return False
            Dim enabledProperty As String = "enabled"
            Dim idIdentity = Me.Item.getProperty("enabled_identity", "")
            If idIdentity <> "" Then
                'Note: If no idIdentiry is set the standard enabled will be used
                If CCO.Permissions.IdentityListHasId(Security.Permissions.Current.IdentitiesList, idIdentity) Then
                    ' User is a part of the identity
                    enabledProperty = "enabled_for_identity"
                End If
            End If
            Dim enabledVal As String = Item.getProperty(enabledProperty, "1")
            Dim res As Boolean = If(enabledVal = "1", True, False)
            Return res
        End Get
    End Property

    Public ReadOnly Property Value As String
        Get
            If Item.isError Then Return "N/A"
            Return Item.getProperty("value", "")
        End Get
    End Property

    Private Function FetchItem() As Item
        If String.IsNullOrEmpty(Me.PropertyName) Then Return Innovator.newError("No name specified for properety")
        Dim amlQuery As String = $"<AML><Item action='get' type='MY Properties' >
                <name condition='eq'>{Me.PropertyName}</name>
            </Item></AML>"
        Dim result = Me.Innovator.applyAML(amlQuery.ToString)
        Return result
    End Function

End Class
