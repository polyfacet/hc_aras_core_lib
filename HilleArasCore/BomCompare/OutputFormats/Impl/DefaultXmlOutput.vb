Imports Aras.IOM
Imports System.Text
Imports System.Xml

Namespace BomCompare.OutputFormat

    Public Class DefaultXmlOutput
        Implements IXmlOutput

        Private Const ITEM_NUMBER As String = "item_number"
        Private Const MAJOR_REV As String = "major_rev"
        Private Const GENERATION As String = "generation"

        Private Property Wr As XmlWriter
        Private Property BomCompareProperties As List(Of IBomCompareItemProperty)
        Public Property OutputSettings As IOutputSettings Implements IXmlOutput.OutputSettings

        Public ReadOnly Property Name As String Implements IXmlOutput.Name
            Get
                Return "Default"
            End Get
        End Property

        'Public Sub New(ByVal outPutSettings As IOutputSettings)
        '    Me.OutPutSettings = outPutSettings
        'End Sub

        Public Function GetResult(ByVal compareItem As Item, ByVal baseItem As Item, ByVal bomCompareProperties As List(Of IBomCompareItemProperty), ByVal bomCompareRows As List(Of BomCompareRow)) As XmlDocument Implements IXmlOutput.GetResult
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
                For Each row As BomCompareRow In bomCompareRows
                    WriteRow(row)
                Next
                Wr.WriteEndElement()

                Wr.WriteEndElement()

            End Using


            Return xmlDoc
        End Function

        Private Sub WriteRow(row As BomCompareRow)
            Wr.WriteStartElement("tr")
            Wr.WriteAttributeString("class", row.ChangeType.ToString)

            Wr.WriteStartElement("td")
            Wr.WriteAttributeString("name", "ChangeType")
            SetCellBgColor(row.ChangeType, False)
            Wr.WriteString(row.ChangeType.ToString)
            Wr.WriteEndElement()

            ' Write Compare Item
            'Wr.WriteStartElement("CompareRowItem")
            WriteItem(row.CompareItem, row.ChangeType)
            'Wr.WriteEndElement()

            Wr.WriteStartElement("td")
            Wr.WriteAttributeString("class", "splitter")
            Wr.WriteAttributeString("bgColor", Me.OutPutSettings.CellColors.SplitterCellColor)
            Wr.WriteString("--")
            Wr.WriteEndElement()

            ' Write Base Item
            'Wr.WriteStartElement("BaseRowItem")
            WriteItem(row.BaseItem, row.ChangeType)
            'Wr.WriteEndElement()


            Wr.WriteStartElement("td")
            Dim i As Integer = 0
            For Each desc As String In row.ChangeDescription
                If i > 0 Then
                    ' Add row break
                    'Wr.WriteElementString("br", "")
                    desc = " , " & desc
                Else
                    If row.ChangeType = BomCompareRow.ChangeTypeEnum.Updated AndAlso Me.OutPutSettings.CellColors.ChangedRowColor = String.Empty Then
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
                    If prop.PropertyName = "item_number" Then
                        Dim linkValue As String = String.Format("'{0}','{1}'", rowItem.TypeName, rowItem.Id)
                        Wr.WriteAttributeString("link", linkValue)
                    End If
                    Wr.WriteString(prop.Value)
                    Wr.WriteEndElement()
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

        Private Sub WriteHeaderItem(ByVal name As String, ByVal item As Item)
            Wr.WriteStartElement(name)
            Wr.WriteElementString("id", item.getProperty("id"))
            Wr.WriteElementString(ITEM_NUMBER, item.getProperty(ITEM_NUMBER))
            Wr.WriteElementString(MAJOR_REV, item.getProperty(MAJOR_REV))
            Wr.WriteElementString(GENERATION, item.getProperty(GENERATION))
            Wr.WriteEndElement()
        End Sub

        Private Sub SetCellBgColor(ByVal changeType As BomCompareRow.ChangeTypeEnum, ByVal row As Boolean)
            Dim valueColor As String = String.Empty

            Select Case changeType
                Case BomCompareRow.ChangeTypeEnum.Added
                    valueColor = Me.OutPutSettings.CellColors.AddedRowColor
                Case BomCompareRow.ChangeTypeEnum.Removed
                    valueColor = Me.OutPutSettings.CellColors.RemovedRowColor
                Case BomCompareRow.ChangeTypeEnum.Replaced
                    valueColor = Me.OutPutSettings.CellColors.ReplacedRowColor
                Case BomCompareRow.ChangeTypeEnum.Updated
                    valueColor = Me.OutPutSettings.CellColors.ChangedRowColor
            End Select

            If Not row Then
                ' If change 
                If changeType = BomCompareRow.ChangeTypeEnum.Updated Then
                    valueColor = Me.OutPutSettings.CellColors.ChangedCellColor
                End If
            End If

            If Not valueColor = String.Empty Then
                Wr.WriteAttributeString("bgColor", valueColor)
            End If

        End Sub

    End Class

End Namespace
