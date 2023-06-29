Imports Aras.IOM
Imports System.Xml
Imports System.Text
Imports System.IO

Namespace Files
    Public Class File
        Inherits InnovatorBase

        'Private Const VAULT_SERVER_CONFIG_FILE_NAME_R11 = "VaultServerConfig.xml"
        Private Const VAULT_SERVER_CONFIG_FILE_NAME = "vault.config"

        Public Sub New(innovator As Innovator)
            MyBase.New(innovator)
        End Sub

        ''' <summary>
        ''' Copies file from vault
        ''' </summary>
        ''' <param name="fileItem"></param>
        ''' <param name="outputDir"></param>
        ''' <returns>Full file path if successful, else empty string</returns>
        Public Function CopyFile(fileItem As Item, outputDir As String) As String
            Dim localVaultPath As String = GetLocalVaultPath()
            Dim databaseName As String = Innovator.getConnection.GetDatabaseName
            Dim baseDirPath As String = localVaultPath & databaseName & "\"

            Dim treePathForItem As String = GetTreePathForFileItem(fileItem)

            Dim fileName As String = fileItem.getProperty("filename")

            Dim sourcePath As String = baseDirPath & treePathForItem
            Dim sourceDir As String = sourcePath
            sourcePath = IO.Path.Combine(sourcePath, fileName)

            Dim destinationPath As String = outputDir
            If Not destinationPath.EndsWith("\") Then
                destinationPath &= "\"
            End If
            destinationPath = IO.Path.Combine(destinationPath, fileName)

            If Not IO.File.Exists(sourcePath) Then
                ' File not in vault as fileName
                ' Check folder for file
                Dim files() As String = IO.Directory.GetFiles(sourceDir)
                If files.Length > 0 Then
                    ' Retrun first, should only be one
                    Dim fileName2 As String = IO.Path.GetFileName(files(0))
                    sourcePath = IO.Path.Combine(sourceDir, fileName2)
                Else
                    ' File not found in vault
                    Return String.Empty
                End If
            End If

            If IO.File.Exists(destinationPath) Then
                IO.File.Delete(destinationPath)
            End If

            IO.File.Copy(sourcePath, destinationPath, True)
            Return destinationPath
        End Function

        Private Function GetTreePathForFileItem(fileItem As Item) As String
            ' Example output result: \4\4C\9E417CF5B454FB978A4BCCEEDC972\
            Dim sb As New StringBuilder
            Dim id As String = fileItem.getID
            sb.Append(id.Chars(0))
            sb.Append("\")
            sb.Append(id.Chars(1))
            sb.Append(id.Chars(2))
            sb.Append("\")
            sb.Append(id.Substring(3))
            sb.Append("\")
            Return sb.ToString
        End Function

        Private Function GetLocalVaultPath() As String
            Dim applicationPath As String = AppDomain.CurrentDomain.BaseDirectory
            Dim di As DirectoryInfo = New DirectoryInfo(applicationPath)
            ' R11 Dim dir As String = di.Parent.Parent.FullName
            Dim dir As String = Path.Combine(di.Parent.Parent.FullName, "VaultServer/")
            Dim fileName As String = VAULT_SERVER_CONFIG_FILE_NAME
            Dim filePath As String = Path.Combine(dir, fileName)

            Dim xmlDoc As New XmlDocument
            xmlDoc.Load(filePath)
            Dim pathNode As XmlNode = xmlDoc.SelectSingleNode("//appSettings/add[@key='LocalPath']")
            Return pathNode.Attributes.GetNamedItem("value").Value
            ' R11 Return xmlDoc.SelectSingleNode("//LocalPath").InnerText
        End Function
    End Class

End Namespace