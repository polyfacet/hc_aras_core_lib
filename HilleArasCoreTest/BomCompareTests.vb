Option Strict On

Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Aras.IOM
Imports Hille.Aras.Core
Imports System.Xml

<TestClass>
Public Class BomCompareTests

    Private Const PART_ITEM_NUMBER = "1020"

    Private innField As Innovator
    Private ReadOnly Property Inn As Innovator
        Get
            If innField Is Nothing Then
                innField = New Connection().GetInnovator
            End If
            Return innField
        End Get
    End Property


    'Private Inn As Innovator

    <TestMethod>
    Public Sub DefaultBomCompare()
        Dim PartNo As String = PART_ITEM_NUMBER

        Inn.getConnection() ' Bara för att trigga inloggning, så att inte tidsmätning påverkas av initial inloggning

        'Console.WriteLine("Connect to Aras: " & InnovatorBase.ToStringAddress)
        'Inn = InnovatorBase.getInnovator
        Dim timer As New Utils.Timer.TaskTimer

        Dim uid As String = Inn.getUserID
        Console.WriteLine("Retrived user id in (ms) " & timer.Restart)

        Dim part As Item = GetPart(Inn, PartNo)
        Console.WriteLine("Retrived part in (ms) " & timer.Restart)
        Console.WriteLine("Generation: " & part.getProperty("generation"))

        Dim previousReleasedRevision As Item = GetPreviousReleasedItem(part)
        Console.WriteLine("Previous recieved in (ms) " & timer.Restart)
        Dim previousRevisionId As String = previousReleasedRevision.getID
        Console.WriteLine("Generation compare: " & previousReleasedRevision.getProperty("generation"))

        Dim comp As New BomCompare.PartBomCompare(Inn)
        Dim compareProperty As BomCompare.IBomCompareItemProperty
        compareProperty = New BomCompare.RelationshipProperty("sort_order", "Seq")

        Dim compProps As New List(Of BomCompare.IBomCompareItemProperty)
        compProps.Add(compareProperty)
        compProps.Add(New BomCompare.NonRelProperty("major_rev", "Revision"))
        compProps.Add(New BomCompare.NonRelProperty("item_number", "Number"))
        compProps.Add(New BomCompare.RelationshipProperty("quantity", "Quantity"))
        comp.SetProperties(compProps)

        Dim result As List(Of String) = comp.Compare(part.getID, previousRevisionId, compareProperty)
        Console.WriteLine("Compared in (ms) " & timer.Restart)
        Console.WriteLine("Rows: " & comp.BomCompareRows.Count)
        Console.WriteLine("Tot time actual work: " & timer.GetTotalTimeInMs)
        For Each row As BomCompare.BomCompareRow In comp.BomCompareRows
            Console.WriteLine(row.ChangeType.ToString)
            For Each desc As String In row.ChangeDescription
                Console.WriteLine(desc)
            Next
        Next
    End Sub


    <TestMethod>
    Public Sub PartDocumentCompare()
        Dim PartNo As String = PART_ITEM_NUMBER
        Dim part As Item = GetPart(Inn, PartNo)
        Console.WriteLine("Generation: " & part.getProperty("generation"))

        Dim previousReleasedRevision As Item = GetPreviousReleasedItem(part)
        Dim previousRevisionId As String = previousReleasedRevision.getID

        Dim comp As New BomCompare.PartBomCompare(Inn)
        comp.RelationshipName = "Part Document"

        Dim compareProperty As BomCompare.IBomCompareItemProperty
        compareProperty = New BomCompare.NonRelProperty("item_number", "Document number")

        Dim compProps As New List(Of BomCompare.IBomCompareItemProperty)
        compProps.Add(compareProperty)
        compProps.Add(New BomCompare.NonRelProperty("major_rev", "Revision"))
        comp.SetProperties(compProps)

        Dim result As List(Of String) = comp.Compare(part.getID, previousRevisionId, compareProperty)
        Console.WriteLine("Rows: " & comp.BomCompareRows.Count)
        For Each row As BomCompare.BomCompareRow In comp.BomCompareRows
            Console.WriteLine(row.ChangeType.ToString)
            For Each desc As String In row.ChangeDescription
                If row.ChangeType <> Hille.Aras.Core.BomCompare.BomCompareRow.ChangeTypeEnum.Unchanged Then
                    For Each prop As BomCompare.IBomCompareItemProperty In row.CompareItem.BomCompareItemProperties
                        Console.Write(prop.PropertyLabel & "=" & prop.Value & ", ")
                    Next
                    Console.WriteLine(desc)
                End If
            Next
        Next
    End Sub

    Private Function GetPart(inn As Innovator, partNo As String) As Item
        Dim aml As String = "<AML><Item action='get' type='Part'><item_number>{0}</item_number></Item></AML>"
        aml = String.Format(aml, partNo)
        Return inn.applyAML(aml)
    End Function

    ' Gets previous revision
    Private Function GetPreviousReleasedItem(ByVal releasedPart As Item) As Item
        Dim previousRev As Item = Nothing
        Dim config_id As String = releasedPart.getProperty("config_id")
        ' Ask for all generations within a revision
        Dim allReleasedGenerations As Item = Inn.newItem(releasedPart.getType, "get")
        allReleasedGenerations.setAttribute("orderBy", "generation DESC")
        allReleasedGenerations.setProperty("config_id", config_id)
        allReleasedGenerations.setProperty("is_released", "1")
        allReleasedGenerations.setPropertyAttribute("id", "condition", "is not null")
        allReleasedGenerations = allReleasedGenerations.apply
        Dim i As Integer = 0
        Do While (i < allReleasedGenerations.getItemCount)
            Dim generationItem As Item = allReleasedGenerations.getItemByIndex(i)
            If (generationItem.getID <> releasedPart.getID) Then
                ' Previous generation
                Return generationItem
            End If
            i = (i + 1)
        Loop

        Return previousRev
    End Function



    <TestMethod>
    Public Sub BomCompare()

        Dim ITEM_NUMBER As String = PART_ITEM_NUMBER

        Dim part As Item = GetPart(Inn, ITEM_NUMBER)
        Dim currentId As String = part.getID
        Dim previousReleasedRevision As Item = GetPreviousReleasedItem(part)
        Dim idPreviousReleased As String = previousReleasedRevision.getID

        Dim indexCompareProperty As BomCompare.IBomCompareItemProperty
        indexCompareProperty = New BomCompare.RelationshipProperty("sort_order", "Sequence")

        Dim bomCompare As New BomCompare.PartBomCompare(Inn)
        Dim compProperties As List(Of BomCompare.IBomCompareItemProperty) = CompareProperties(indexCompareProperty)
        bomCompare.SetProperties(compProperties)

        bomCompare.ChangeDescription = " this value {0} compared value {1}"

        bomCompare.Compare(currentId, idPreviousReleased, indexCompareProperty)

        For Each row As BomCompare.BomCompareRow In bomCompare.BomCompareRows
            Console.WriteLine(row.ChangeType.ToString)
        Next

        ' Get output
        Dim xmlDoc As Xml.XmlDocument
        Dim outputSettings As BomCompare.OutputFormat.IOutputSettings = New BomCompare.OutputFormat.DefaultOutputSettings
        outputSettings.CellColors.ChangedRowColor = "lightyellow"
        outputSettings.CellColors.ChangedCellColor = "yellow"
        xmlDoc = bomCompare.GetResultAsXml(outputSettings)

        ' Convert xml to html and display
        Dim html As String = ConvertToHtml(xmlDoc)
        Dim tempFilePath As String = IO.Path.Combine(IO.Path.GetTempPath, Date.Now.Ticks.ToString & ".html")
        Using sw As New IO.StreamWriter(tempFilePath)
            sw.WriteLine(html)
        End Using

        Process.Start(tempFilePath)
        Threading.Thread.Sleep(5000) ' Wait five sec then delete file
        IO.File.Delete(tempFilePath)

        Console.WriteLine("Done")


    End Sub

    Private Function ConvertToHtml(xmlDoc As XmlDocument) As String
        Dim sb As New Text.StringBuilder
        sb.AppendLine("<!DOCTYPE html>")
        sb.AppendLine("<meta charset='UTF-8'>")
        sb.AppendLine("<html>")

        sb.AppendLine(xmlDoc.GetElementsByTagName("table").Item(0).OuterXml)

        sb.AppendLine("</html>")
        Return sb.ToString
    End Function

    Private Function CompareProperties(indexProp As BomCompare.IBomCompareItemProperty) As List(Of BomCompare.IBomCompareItemProperty)
        Dim list As New List(Of BomCompare.IBomCompareItemProperty)

        list.Add(indexProp)

        ' Include but dont describe comparision on name
        Dim nameProp As New BomCompare.NonRelProperty("name", "Name")
        nameProp.IsComparable = False

        list.Add(nameProp)
        list.Add(New BomCompare.RelationshipProperty("quantity", "Quantity"))
        list.Add(New BomCompare.NonRelProperty("item_number", "Number"))


        Return list
    End Function

End Class
