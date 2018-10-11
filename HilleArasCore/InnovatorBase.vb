Imports System.IO
Imports Aras.IOM

Public MustInherit Class InnovatorBase

    Public Sub New(ByVal innovator As Innovator)
        Me.Innovator = innovator
    End Sub

    Protected ReadOnly Property Innovator As Innovator

    Public Shared Function GetArasTempDir()
        Dim applicationPath As String = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath
        Dim di As DirectoryInfo = New DirectoryInfo(applicationPath)
        Dim dir As String = di.Parent.FullName
        dir = IO.Path.Combine(dir, "\Server\temp")
        Return dir
    End Function

End Class
