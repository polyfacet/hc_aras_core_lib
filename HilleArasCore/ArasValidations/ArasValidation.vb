Imports Aras.IOM

Public Class ArasValidation
    Private Const MY_VALIDATION_TYPE As String = "MY_Validation"

    Public Enum ValidationType
        [Error]
        Warning
        Info
        Ok
    End Enum

    Private Const BASE_PATH As String = "..\Solutions\HilleConsultIt\images\"

    Private ImagePathField As String = String.Empty

    Public Sub New(validationType As ValidationType, item As Item, description As String)
        Me.Type = validationType
        Me.Item = item
        Me.Description = description
        If Not item Is Nothing Then
            Me.ItemNumber = item.getProperty("item_number", "")
        End If
    End Sub

    Public Sub New(validationType As ValidationType, description As String)
        Me.Type = validationType
        Me.Description = description
    End Sub

    Public Property Type As ValidationType
    Public Property Description As String
    Public Property ItemNumber As String
    Public Property LongDescription As String
    Public Property Item As Item

    Public ReadOnly Property ImagePath As String
        Get
            ImagePathField = BASE_PATH + Me.Type.ToString() + ".png"
            Return ImagePathField
        End Get
    End Property

    Public Shared Function ConvertValidationToItemResult(inn As Innovator, validations As List(Of ArasValidation), sourceId As String, validationItemType As String) As Item
        Dim res As Item = inn.newItem("Dummy")
        Dim validationItems As Item = Nothing
        For Each validation As ArasValidation In validations
            Dim validationItem As Item = ConvertValidationToItem(inn, validation)
            If validationItems Is Nothing Then
                validationItems = validationItem
            Else
                validationItems.appendItem(validationItem)
            End If
            Dim rel As Item = inn.newItem(validationItemType)
            rel.setID(inn.getNewID())
            rel.setAttribute("typeId", "dummy")
            rel.setPropertyItem("related_id", validationItem)
            rel.removeAttribute("isNew")
            rel.removeAttribute("isTemp")
            rel.setProperty("item_number", validation.ItemNumber)
            If Not validation.Item Is Nothing Then
                rel.setProperty("item", validation.Item.getID())
                rel.setPropertyAttribute("item", "keyed_name", validation.ItemNumber)
            End If
            rel.setProperty("source_id", sourceId)
            res.appendItem(rel)
        Next

        If res.getItemCount() > 1 Then
            res.removeItem(res.getItemByIndex(0))
            Dim res1 As Item = inn.newResult("")
            res1.dom.SelectSingleNode("//Result").InnerXml = res.dom.SelectSingleNode("//AML").InnerXml
            Return res1
        Else
            Return res
        End If
    End Function

    Private Shared Function ConvertValidationToItem(inn As Innovator, validation As ArasValidation) As Item
        Dim validationItem As Item = inn.newItem(MY_VALIDATION_TYPE, "add")
        validationItem.setProperty("validation_type", validation.Type.ToString())
        validationItem.setProperty("image", validation.ImagePath)
        validationItem.setProperty("description", validation.Description)
        validationItem.setProperty("long_description", validation.LongDescription)
        validationItem.setProperty("item_number", validation.ItemNumber)
        If Not validation.Item Is Nothing Then
            validationItem.setProperty("item_type", validation.Item.getType())
        End If
        validationItem.removeAttribute("isNew")
        validationItem.removeAttribute("isTemp")
        validationItem.apply()
        Return validationItem
    End Function


End Class
