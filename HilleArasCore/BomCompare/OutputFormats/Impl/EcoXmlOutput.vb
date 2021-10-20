Imports System.Text
Imports System.Xml
Imports Aras.IOM
Imports Hille.Aras.Core
Imports Hille.Aras.Core.BomCompare
Imports Hille.Aras.Core.BomCompare.OutputFormat

Public Class EcoXmlOutput
    Implements IXmlOutput

    Private Const ITEM_NUMBER As String = "item_number"
    Private Const MAJOR_REV As String = "major_rev"
    Private Const GENERATION As String = "generation"
    Private Const MY_DESCRIPTION As String = "my_description"


    Private Property Wr As XmlWriter
    Private Property BomCompareProperties As List(Of IBomCompareItemProperty)

    Public ReadOnly Property Name As String Implements IXmlOutput.Name
        Get
            Return "ECO"
        End Get
    End Property

    Public Property OutputSettings As IOutputSettings Implements IXmlOutput.OutputSettings

    Public Function GetResult(compareItem As Item, baseItem As Item, bomCompareProperties As List(Of IBomCompareItemProperty), bomCompareRows As List(Of BomCompareRow)) As XmlDocument Implements IXmlOutput.GetResult
        Me.BomCompareProperties = bomCompareProperties
        Dim settings As XmlWriterSettings = New XmlWriterSettings()
        settings.Indent = True

        Dim xmlSb As New StringBuilder()

        Dim xmlDoc As New XmlDocument()
        Me.Wr = xmlDoc.CreateNavigator.AppendChild
        Using Me.Wr

            Wr.WriteStartElement("BomCompareResult")

            Me.WriteHeaderItem("CompareItem", compareItem)
            Me.WriteHeaderItem("BaseItem", baseItem)

            Wr.WriteStartElement("table")

            If bomCompareRows.Count > 0 Then
                WriteTableHeader(bomCompareRows.Item(0))
            End If


            For Each row As BomCompareRow In bomCompareRows
                Dim changeType As BomCompareRow.ChangeTypeEnum = row.ChangeType
                Select Case changeType
                    Case BomCompareRow.ChangeTypeEnum.Unchanged
                        ' Skip unchanged rows
                    Case Else
                        WriteRow(row)
                End Select

            Next
            Wr.WriteEndElement()

            Wr.WriteEndElement()

        End Using


        Return xmlDoc
    End Function

    Private Sub WriteTableHeader(row As BomCompareRow)
        Wr.WriteStartElement("tr")
        Wr.WriteAttributeString("class", "ChangesHeader")

        ' Change type header
        Wr.WriteStartElement("th")
        Wr.WriteAttributeString("name", "ChangeState")
        Wr.WriteString("State")
        Wr.WriteEndElement()

        Dim rowItem As BomCompareItem = If(Not row.CompareItem Is Nothing, row.CompareItem, row.BaseItem)
        For Each prop As IBomCompareItemProperty In rowItem.BomCompareItemProperties
            Dim propName As String = If(String.IsNullOrEmpty(prop.PropertyName), "", prop.PropertyName)
            If Not (propName.ToUpper.Contains("MSEQ") Or propName.ToUpper.Contains("POS") Or propName.ToUpper.Equals("ID")) Then ' Don't include pos or mseq or ID
                Wr.WriteStartElement("th")
                Wr.WriteAttributeString("name", prop.PropertyLabel)
                Wr.WriteString(prop.PropertyLabel)
                Wr.WriteEndElement()
            End If
        Next

        ' Change description
        Wr.WriteStartElement("th")
        Wr.WriteAttributeString("name", "Description")
        Wr.WriteString("Change description")
        Wr.WriteEndElement()

        Wr.WriteEndElement()
    End Sub

    Private Sub WriteHeaderItem(ByVal name As String, ByVal item As Item)
        Wr.WriteStartElement(name)
        Wr.WriteElementString("id", item.getProperty("id"))
        Wr.WriteElementString(ITEM_NUMBER, item.getProperty(ITEM_NUMBER))
        Wr.WriteElementString(MAJOR_REV, item.getProperty(MAJOR_REV))
        Wr.WriteElementString(GENERATION, item.getProperty(GENERATION))
        Wr.WriteElementString(MY_DESCRIPTION, item.getProperty(MY_DESCRIPTION))
        Wr.WriteEndElement()
    End Sub

    Private Sub WriteRow(row As BomCompareRow)
        Wr.WriteStartElement("tr")
        Wr.WriteAttributeString("class", row.ChangeType.ToString)

        Wr.WriteStartElement("td")
        Wr.WriteAttributeString("name", "ChangeType")
        SetCellBgColor(row.ChangeType, False)
        Wr.WriteString(row.ChangeType.ToString)
        Wr.WriteEndElement()

        ' Write Compare Item
        Dim rowItem As BomCompareItem = If(Not row.CompareItem Is Nothing, row.CompareItem, row.BaseItem)
        WriteItem(rowItem, row.ChangeType)

        Wr.WriteStartElement("td")
        Dim i As Integer = 0
        For Each desc As String In row.ChangeDescription
            If i > 0 Then
                ' Add row break
                'Wr.WriteElementString("br", "")
                desc = " , " & desc
            Else
                If row.ChangeType = BomCompareRow.ChangeTypeEnum.Updated AndAlso Me.OutputSettings.CellColors.ChangedRowColor = String.Empty Then
                    ' No row color for updated row, set on cell level
                    SetCellBgColor(row.ChangeType, False)
                Else
                    SetCellBgColor(row.ChangeType, True)
                End If
            End If
            Wr.WriteString(desc)
            i += 1
        Next
        Wr.WriteEndElement()

        Wr.WriteEndElement()
    End Sub

    Private Sub WriteItem(rowItem As BomCompareItem, ByVal changeType As BomCompareRow.ChangeTypeEnum)
        If Not rowItem Is Nothing Then
            For Each prop As IBomCompareItemProperty In rowItem.BomCompareItemProperties
                Dim propName As String = If(String.IsNullOrEmpty(prop.PropertyName), "", prop.PropertyName)
                If Not (propName.ToUpper.Contains("MSEQ") Or propName.ToUpper.Contains("POS") Or propName.ToUpper.Equals("ID")) Then ' Don't include pos or mseq or Id
                    Wr.WriteStartElement("td")
                    Wr.WriteAttributeString("name", prop.PropertyLabel)
                    If prop.IsComparable AndAlso prop.Changed Then
                        Wr.WriteAttributeString("changed", prop.Changed.ToString)
                        ' Set cell specific color
                        SetCellBgColor(changeType, False)
                    Else
                        SetCellBgColor(changeType, True)
                    End If
                    'Wr.WriteElementString(prop.PropertyName, prop.Value)
                    Wr.WriteString(prop.Value)
                    Wr.WriteEndElement()
                End If
            Next
        Else
            ' Add empty cells
            For Each prop As IBomCompareItemProperty In Me.BomCompareProperties
                Wr.WriteStartElement("td")
                SetCellBgColor(changeType, True)
                Wr.WriteEndElement()
            Next

        End If
    End Sub

    Private Sub SetCellBgColor(ByVal changeType As BomCompareRow.ChangeTypeEnum, ByVal row As Boolean)
        Dim valueColor As String = String.Empty

        Select Case changeType
            Case BomCompareRow.ChangeTypeEnum.Added
                valueColor = Me.OutputSettings.CellColors.AddedRowColor
            Case BomCompareRow.ChangeTypeEnum.Removed
                valueColor = Me.OutputSettings.CellColors.RemovedRowColor
            Case BomCompareRow.ChangeTypeEnum.Replaced
                valueColor = Me.OutputSettings.CellColors.ReplacedRowColor
            Case BomCompareRow.ChangeTypeEnum.Updated
                valueColor = Me.OutputSettings.CellColors.ChangedRowColor
        End Select

        If Not row Then
            ' If change 
            If changeType = BomCompareRow.ChangeTypeEnum.Updated Then
                valueColor = Me.OutputSettings.CellColors.ChangedCellColor
            End If
        End If

        If Not valueColor = String.Empty Then
            Wr.WriteAttributeString("bgColor", valueColor)
        End If

    End Sub

End Class
