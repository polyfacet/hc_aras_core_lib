Imports Aras.IOM
Imports System.Xml
Imports Hille.Aras.Core.BomCompare
Imports Hille.Aras.Core.BomCompare.OutputFormat

Public Interface IXmlOutput

    ReadOnly Property Name As String
    Property OutputSettings As IOutputSettings
    Function GetResult(ByVal compareItem As Item, ByVal baseItem As Item, ByVal bomCompareProperties As List(Of IBomCompareItemProperty), ByVal bomCompareRows As List(Of BomCompareRow)) As XmlDocument


End Interface
