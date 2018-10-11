Imports Aras.IOM
Imports System.Text

Public Class DefaultArasExceptionLogger
    Implements IArasExceptionLogger

    Private Const DATE_FORMAT As String = "yyyy-MM-dd HH:mm:ss"


    Public Sub New(inn As Innovator, classification As String)
        Me.Inn = inn
        SB = New StringBuilder
        Me.Classification = classification
    End Sub

    Private Property Inn As Innovator

    Protected Overridable Property Classification As String

    Private logItemField As Item
    Private Property LogItem As Item
        Get
            If logItemField Is Nothing Then
                If Not Inn Is Nothing Then
                    ' Create new logItem
                    logItemField = Inn.newItem("MY Log", "add")
                End If
            End If
            Return logItemField
        End Get
        Set(value As Item)
            logItemField = value
        End Set
    End Property

    Private Property SB As StringBuilder


    Private logDirField As String
    Private Property LogDir As String
        Get
            If String.IsNullOrEmpty(logDirField) Then
                logDirField = GetLogDir(Me.Inn)
            End If
            Return logDirField
        End Get
        Set(value As String)
            logDirField = value
        End Set
    End Property

    Private Shared Function GetLogDir(ByVal inn As Innovator) As String
        Dim logDir As String = GetBaseLogDir()

        If Not inn Is Nothing Then
            logDir &= "\" & inn.getConnection.GetDatabaseName()
        End If

        If Not IO.Directory.Exists(logDir) Then
            IO.Directory.CreateDirectory(logDir)
        End If
        Return logDir
    End Function

    Private Shared Function GetBaseLogDir() As String
        Dim logDir As String = InnovatorBase.GetArasTempDir

        logDir = IO.Path.Combine(logDir, "ArasExceptions")
        Return logDir
    End Function

    Public Sub Log(inn As Innovator, method As String, message As String, innerException As Exception) Implements IArasExceptionLogger.Log
        Dim fileLogMessage As String
        Me.Inn = inn

        Dim userId As String = "N/A"
        Dim userLogin As String = "N/A"

        SB.AppendLine("*********START Exception*********")
        SB.AppendLine(DateTime.Now.ToString(DATE_FORMAT))

        If Not inn Is Nothing Then

            AppendProperty("classification", Classification)

            userId = inn.getUserID()
            Dim user As Item = inn.getItemById("User", userId)
            userLogin = user.getProperty("login_name")

        End If

        SB.AppendLine("UserId: " & userId)
        AppendProperty("user_login", userLogin)

        AppendProperty("method", method)

        AppendProperty("message", message)


        If Not innerException Is Nothing Then
            AppendProperty("inner_exception", DateTime.Now.ToString(DATE_FORMAT) & vbCrLf & innerException.ToString)
        Else
            AppendProperty("inner_exception", DateTime.Now.ToString(DATE_FORMAT))
        End If

        SB.AppendLine("*********END Exception*********")
        fileLogMessage = SB.ToString

        LogToFile(inn, fileLogMessage)
        If Not LogItem Is Nothing Then
            LogToXmlFile(inn, LogItem)
        End If


    End Sub

    Public Shared Sub ImportLogs(ByVal inn As Innovator)
        Dim dir As IO.DirectoryInfo = New IO.DirectoryInfo(GetLogDir(inn))
        Dim xmlFiles() As IO.FileInfo = dir.GetFiles("*.xml")

        For Each xmlFile As IO.FileInfo In xmlFiles
            Dim sb As New Text.StringBuilder
            sb.AppendLine("<AML>")

            Dim fileContent As String = IO.File.ReadAllText(xmlFile.FullName)
            fileContent = fileContent.Replace("[\x01-\x1F]+", "")
            sb.AppendLine(fileContent)


            sb.AppendLine("</AML>")
            Dim result As Item = inn.applyAML(sb.ToString)
            If Not result.isError Then
                xmlFile.Delete()
            End If
        Next


    End Sub

    Protected Sub AppendProperty(propertyName As String, propertyValue As String)
        SB.AppendLine(propertyName & ": " & propertyValue)
        If Not LogItem Is Nothing Then
            LogItem.setProperty(propertyName, propertyValue)
        End If
    End Sub


    Private Sub LogToFile(inn As Innovator, logMessage As String)
        Dim logDir As String = GetLogDir(inn)

        If Not IO.Directory.Exists(logDir) Then
            IO.Directory.CreateDirectory(logDir)
        End If
        Dim fileName As String = DateTime.Now.Ticks.ToString & ".log"
        Dim filePath As String = IO.Path.Combine(logDir, fileName)

        Using sw As New IO.StreamWriter(filePath, True)
            sw.WriteLine(logMessage)
        End Using
    End Sub

    Private Sub LogToXmlFile(inn As Innovator, ByVal item As Item)
        Dim id As String = item.dom.GetElementsByTagName("Item")(0).Attributes.GetNamedItem("id").Value
        Dim filePath As String = IO.Path.Combine(LogDir, id & ".xml")
        item.dom.Save(filePath)
    End Sub

End Class
