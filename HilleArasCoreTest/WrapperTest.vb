Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Aras.IOM
Imports Hille.Aras.Core

<TestClass>
Public Class WrapperTests

    Private innField As Innovator
    Private ReadOnly Property Inn As Innovator
        Get
            If innField Is Nothing Then
                innField = InnovatorBase.getInnovator
            End If
            Return innField
        End Get
    End Property

    <TestMethod>
    Public Sub WrapTest1()

        Dim ITEM_NUMBER As String = "1045"
        Dim amlQueryPart As String = "<AML><Item action='get' type='Part'><item_number>{0}</item_number></Item></AML>"
        amlQueryPart = String.Format(amlQueryPart, ITEM_NUMBER)

        Dim part As Item = Inn.applyAML(amlQueryPart)

        Hille.Aras.Core.ArasMethodWrapper.RunArasMethod(part, AddressOf MainTest)

        Hille.Aras.Core.ArasMethodWrapper.RunArasMethod(part, Function(itemMethod As Item) As Item
                                                                  Dim inn As Innovator = itemMethod.getInnovator
                                                                  Console.WriteLine("Sec test")
                                                                  Threading.Thread.Sleep(50)
                                                                  Return itemMethod
                                                              End Function
                                                                  )

    End Sub

    Function MainTest(itemMethod As Item) As Item
        Dim inn As Innovator = itemMethod.getInnovator()
        Console.WriteLine("MainTest: " & itemMethod.getID)
        Threading.Thread.Sleep(100)
        Return itemMethod
    End Function

End Class
