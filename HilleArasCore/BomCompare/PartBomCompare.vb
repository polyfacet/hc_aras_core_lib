Imports Aras.IOM
Imports System.Text
Imports System.Xml

Namespace BomCompare

    Public Class PartBomCompare
        Inherits InnovatorBase

        Private Property BomCompareProperties As List(Of IBomCompareItemProperty)
        Private Property CompareProperty As IBomCompareItemProperty
        Private Property CompareItem As Item
        Private Property BaseItem As Item

        Private changeDescriptionField As String
        Public Property ChangeDescription As String
            Get
                If String.IsNullOrEmpty(changeDescriptionField) Then
                    changeDescriptionField = " this value {0} compared value {1}"
                End If
                Return changeDescriptionField
            End Get
            Set(value As String)
                changeDescriptionField = value
            End Set
        End Property



        Public Sub New(innovator As Innovator)
            MyBase.New(innovator)
        End Sub

        Public Sub SetProperties(ByVal properties As List(Of IBomCompareItemProperty))
            BomCompareProperties = properties
        End Sub

        Private bomCompareRowsField As New List(Of BomCompareRow)
        Public ReadOnly Property BomCompareRows As List(Of BomCompareRow)
            Get
                Return bomCompareRowsField
            End Get
        End Property

        ' Make it possible to use other item type than the Part realationship. e.g. Document
        Private itemTypeNameField As String = "Part"
        Public Property ItemTypeName As String
            Get
                Return itemTypeNameField
            End Get
            Set(value As String)
                itemTypeNameField = value
            End Set
        End Property

        ' Make it possible to use other relationship than the Part BOM realationship. e.g. Part Document
        Private relationshipNameField As String = "Part BOM"
        Public Property RelationshipName As String
            Get
                Return relationshipNameField
            End Get
            Set(value As String)
                relationshipNameField = value
            End Set
        End Property


        Public Function GetResultAsXml(ByVal outPutSettings As BomCompare.OutputFormat.IOutputSettings) As XmlDocument
            Dim output As New BomCompare.OutputFormat.DefaultXmlOutput()
            output.OutPutSettings = outPutSettings
            Return output.GetResult(Me.CompareItem, Me.BaseItem, Me.BomCompareProperties, Me.BomCompareRows)
        End Function

        Public Function GetResultAsXml(ByVal outPutSettings As BomCompare.OutputFormat.IOutputSettings, ByVal implementationName As String) As XmlDocument

            If String.IsNullOrEmpty(implementationName) Then
                ' Use default
                Return GetResultAsXml(outPutSettings)
            End If

            ' Get specific implementation
            Dim output As IXmlOutput = XmlOutputFactory.GetImplemenation(implementationName)
            output.OutputSettings = outPutSettings
            Return output.GetResult(Me.CompareItem, Me.BaseItem, Me.BomCompareProperties, Me.BomCompareRows)

        End Function

        Public Function GetResultAsXml(ByVal outputImpl As IXmlOutput) As XmlDocument

            ' Get specific implementation 
            Dim output As IXmlOutput = outputImpl
            Return output.GetResult(Me.CompareItem, Me.BaseItem, Me.BomCompareProperties, Me.BomCompareRows)

        End Function


        Public Function Compare(ByVal comparePartId As String, ByVal basePartId As String, ByVal compareProperty As IBomCompareItemProperty) As List(Of String)
            Me.CompareProperty = compareProperty
            Dim sb As New StringBuilder
            sb.Append("<AML>")
            sb.Append("<Item action='get' type='{2}' id='{0}'>")
            sb.Append("<Relationships>")
            sb.Append("<Item action = 'get' type='" & RelationshipName & "' orderBy='{1}'>")
            sb.Append("<is_current condition='eq'>1</is_current>")
            sb.Append("</Item>")
            sb.Append("</Relationships>")
            sb.Append("</Item>")
            sb.Append("</AML>")

            Dim getPartWithBomAml As String = sb.ToString()

            Dim amlQuery As String = String.Format(getPartWithBomAml, comparePartId, Me.CompareProperty.PropertyName, Me.ItemTypeName)
            Dim comparePart As Item = Innovator.applyAML(amlQuery)
            amlQuery = String.Format(getPartWithBomAml, basePartId, Me.CompareProperty.PropertyName, Me.ItemTypeName)
            Dim basePart As Item = Innovator.applyAML(amlQuery)
            Me.CompareItem = comparePart
            Me.BaseItem = basePart
            ' Do compare
            Return Compare(comparePart.dom, basePart.dom)

        End Function

        Private Function Compare(ByVal comparePartDoc As XmlDocument, ByVal basePartDoc As XmlDocument) As List(Of String)
            Dim messages As New List(Of String)
            Dim partBomXpath As String = "//Item[@type='" & RelationshipName & "']"
            Dim partBomCompareNodes As XmlNodeList = comparePartDoc.SelectNodes(partBomXpath)
            Dim partBomBaseNodes As XmlNodeList = basePartDoc.SelectNodes(partBomXpath)



            Dim foundBaseNodes As New Dictionary(Of String, XmlNode)

            For Each partBomNode As XmlNode In partBomCompareNodes
                Dim partBomXmlElement As XmlElement = CType(partBomNode, XmlElement)
                Dim indexValue As String
                Try
                    If TypeOf Me.CompareProperty Is RelationshipProperty Then
                        indexValue = partBomXmlElement.Item(Me.CompareProperty.PropertyName).InnerText
                    Else
                        indexValue = partBomXmlElement.Item("related_id").Item("Item").Item(Me.CompareProperty.PropertyName).InnerText
                    End If
                Catch ex As Exception
                    Throw New ArasException(Me.Innovator, "Missing index value:  " & Me.CompareProperty.PropertyLabel & ", try another index to do compare on.", ex)
                End Try

                Dim partBomBaseXmlElement As XmlElement = FindElementInBaseNodeList(partBomBaseNodes, indexValue)
                Dim bomCompareRow As BomCompareRow

                If partBomBaseXmlElement Is Nothing Then
                    ' New row...
                    Dim itemNumber As String = partBomXmlElement.Item("related_id").GetAttribute("keyed_name")
                    messages.Add("Added: " & itemNumber)
                    bomCompareRow = CreateAddBomCompareRow(indexValue, partBomXmlElement)
                Else
                    ' Found with same index, do a compare on this level for changes in quantity, pos, changed item or other specific changes.
                    Dim diffMessages As New List(Of String)

                    For Each prop As IBomCompareItemProperty In BomCompareProperties
                        If prop.IsComparable AndAlso prop.PropertyName <> Me.CompareProperty.PropertyName Then
                            Dim compareValue As String
                            Dim baseValue As String
                            If prop.IsRelationProperty Then
                                ' Check link property
                                Dim element As XmlElement = partBomXmlElement.Item(prop.PropertyName)
                                compareValue = If((Not element Is Nothing), element.InnerText, String.Empty)
                                element = partBomBaseXmlElement.Item(prop.PropertyName)
                                baseValue = If((Not element Is Nothing), element.InnerText, String.Empty)
                            Else
                                ' Check component
                                compareValue = partBomXmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName).InnerText
                                baseValue = partBomBaseXmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName).InnerText
                            End If

                            If compareValue <> baseValue Then
                                ' Found a diff
                                'diffMessages.Add("Update: " & prop.PropertyLabel & " new value " & compareValue & " old value " & baseValue)
                                diffMessages.Add(prop.PropertyLabel & String.Format(ChangeDescription, compareValue, baseValue))
                            End If

                        End If
                    Next

                    Me.CreateUpdateBomCompareRow(indexValue, diffMessages, partBomXmlElement, partBomBaseXmlElement)

                    ' Check this one as found in the comparePart
                    Dim id As String = partBomBaseXmlElement.GetAttribute("id")
                    If (foundBaseNodes.ContainsKey(id)) Then
                        ' This can be the case when not comparing on realtionship property. E.g. item_number
                        Dim conflictingValue As String
                        If Me.CompareProperty.IsRelationProperty Then
                            Dim element As XmlElement = partBomXmlElement.Item(Me.CompareProperty.PropertyName)
                            conflictingValue = If((Not element Is Nothing), element.InnerText, String.Empty)
                        Else
                            conflictingValue = partBomBaseXmlElement.SelectSingleNode("related_id/Item").Item(Me.CompareProperty.PropertyName).InnerText
                        End If
                        Dim errorMessage As String = "Could not compare by {0} due to duplicate key. Use another compare index. Conflicting value: {1}"
                        errorMessage = String.Format(errorMessage, Me.CompareProperty.PropertyLabel, conflictingValue)
                        Throw New ArasException(Me.Innovator, errorMessage)
                    End If
                    foundBaseNodes.Add(id, partBomBaseXmlElement)

                End If

                'CreateBomCompareRow(indexValue)
            Next

            ' Add the missing relations info
            If foundBaseNodes.Count < partBomBaseNodes.Count Then
                For Each node As XmlElement In partBomBaseNodes
                    Dim thisId As String = node.GetAttribute("id")
                    If Not foundBaseNodes.ContainsKey(thisId) Then
                        Dim index As String
                        If TypeOf Me.CompareProperty Is RelationshipProperty Then
                            index = node.Item(Me.CompareProperty.PropertyName).InnerText
                        Else
                            index = node.SelectSingleNode("related_id/Item").Item(Me.CompareProperty.PropertyName).InnerText
                        End If

                        ' Missing in comparePart
                        Dim itemNumber As String = node.Item("related_id").GetAttribute("keyed_name")
                        Dim removedMessages As New List(Of String)
                        removedMessages.Add("Removed: " & itemNumber)

                        CreateRemovedBomCompareRow(Index, removedMessages, node)
                    End If
                Next
            End If

            Return messages
        End Function

        Private Function FindElementInBaseNodeList(partBomBaseNodes As XmlNodeList, indexValue As String) As XmlElement

            For Each partBomNode As XmlNode In partBomBaseNodes
                Dim partBomXmlElement As XmlElement = CType(partBomNode, XmlElement)
                Dim thisIndexValue As String
                If TypeOf Me.CompareProperty Is RelationshipProperty Then
                    thisIndexValue = partBomXmlElement.Item(Me.CompareProperty.PropertyName).InnerText
                Else
                    thisIndexValue = partBomXmlElement.Item("related_id").Item("Item").Item(Me.CompareProperty.PropertyName).InnerText
                End If

                If thisIndexValue = indexValue Then
                    ' Found match
                    Return partBomXmlElement
                End If
            Next

            Return Nothing
        End Function

        Private Function CreateAddBomCompareRow(ByVal indexValue As String, ByVal xmlElement As XmlElement) As BomCompareRow

            Dim compareItem As New BomCompareItem

            ' Create new instances of the properties
            For Each prop As IBomCompareItemProperty In Me.BomCompareProperties
                If prop.Visible Then
                    Dim newProp As IBomCompareItemProperty
                    Dim propElement As XmlElement
                    If prop.IsRelationProperty Then
                        newProp = New RelationshipProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = xmlElement.Item(prop.PropertyName)

                    Else
                        newProp = New NonRelProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = xmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName)
                    End If

                    If Not propElement Is Nothing Then
                        newProp.Value = propElement.InnerText
                    Else
                        newProp.Value = String.Empty
                    End If

                    compareItem.BomCompareItemProperties.Add(newProp)
                End If
            Next

            Dim description As New List(Of String)
            description.Add("New component added")
            Dim bomCompareRow As BomCompareRow
            bomCompareRow = New BomCompareRow(indexValue, bomCompareRow.ChangeTypeEnum.Added, description, compareItem, Nothing)

            bomCompareRowsField.Add(bomCompareRow)
            Return bomCompareRow
        End Function


        Private Function CreateUpdateBomCompareRow(ByVal indexValue As String, ByVal messages As List(Of String), ByVal compareXmlElement As XmlElement, ByVal baseXmlElement As XmlElement) As BomCompareRow

            Dim compareItem As New BomCompareItem
            Dim baseItem As New BomCompareItem

            ' Create new instances of the properties
            For Each prop As IBomCompareItemProperty In Me.BomCompareProperties
                If prop.Visible Then
                    Dim newProp As IBomCompareItemProperty
                    Dim newPropBase As IBomCompareItemProperty
                    Dim propElement As XmlElement
                    Dim propElementBase As XmlElement
                    If prop.IsRelationProperty Then
                        newProp = New RelationshipProperty(prop.PropertyName, prop.PropertyLabel)
                        newPropBase = New RelationshipProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = compareXmlElement.Item(prop.PropertyName)
                        propElementBase = baseXmlElement.Item(prop.PropertyName)

                    Else
                        newProp = New NonRelProperty(prop.PropertyName, prop.PropertyLabel)
                        newPropBase = New NonRelProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = compareXmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName)
                        propElementBase = baseXmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName)
                    End If

                    If Not propElement Is Nothing Then
                        newProp.Value = propElement.InnerText
                    Else
                        newProp.Value = String.Empty
                    End If

                    If Not propElementBase Is Nothing Then
                        newPropBase.Value = propElementBase.InnerText
                    Else
                        newPropBase.Value = String.Empty
                    End If

                    If newProp.Value <> newPropBase.Value Then
                        newProp.Changed = True
                        newPropBase.Changed = True
                    End If

                    compareItem.BomCompareItemProperties.Add(newProp)
                    baseItem.BomCompareItemProperties.Add(newPropBase)
                End If
            Next


            Dim bomCompareRow As BomCompareRow
            If messages.Count > 0 Then
                bomCompareRow = New BomCompareRow(indexValue, BomCompareRow.ChangeTypeEnum.Updated, messages, compareItem, baseItem)
            Else
                bomCompareRow = New BomCompareRow(indexValue, BomCompareRow.ChangeTypeEnum.Unchanged, messages, compareItem, baseItem)
            End If


            bomCompareRowsField.Add(bomCompareRow)
            Return bomCompareRow


        End Function

        Private Function CreateRemovedBomCompareRow(ByVal indexValue As String, ByVal description As List(Of String), ByVal xmlElement As XmlElement) As BomCompareRow

            Dim baseItem As New BomCompareItem

            ' Create new instances of the properties
            For Each prop As IBomCompareItemProperty In Me.BomCompareProperties
                If prop.Visible Then
                    Dim newProp As IBomCompareItemProperty
                    Dim propElement As XmlElement
                    If prop.IsRelationProperty Then
                        newProp = New RelationshipProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = xmlElement.Item(prop.PropertyName)

                    Else
                        newProp = New NonRelProperty(prop.PropertyName, prop.PropertyLabel)
                        propElement = xmlElement.SelectSingleNode("related_id/Item").Item(prop.PropertyName)
                    End If

                    If Not propElement Is Nothing Then
                        newProp.Value = propElement.InnerText
                    Else
                        newProp.Value = String.Empty
                    End If

                    baseItem.BomCompareItemProperties.Add(newProp)
                End If
            Next

            Dim bomCompareRow As BomCompareRow
            bomCompareRow = New BomCompareRow(indexValue, BomCompareRow.ChangeTypeEnum.Removed, description, Nothing, baseItem)

            bomCompareRowsField.Add(bomCompareRow)
            Return bomCompareRow

        End Function

    End Class



End Namespace