Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Aras.IOM
Imports Hille.Aras.Core

<TestClass>
Public Class CoreTests

    Private Const PART_ITEM_NUMBER = "123"
    Private Const PART_NAME = "Test Part"

    Private innField As Innovator
    Private ReadOnly Property Inn As Innovator
        Get
            If innField Is Nothing Then
                innField = New Connection().GetInnovator
            End If
            Return innField
        End Get
    End Property


    <TestMethod>
    Public Sub ConnectToAras()
        Dim timer As New Hille.Aras.Core.Utils.Timer.TaskTimer
        Dim aml As String = "<AML><Item action='get' type='ItemType'><name>Part</name></Item></AML>"
        Dim partItemType As Item = Inn.applyAML(aml)
        Console.WriteLine("Login time: " & timer.GetElapsedTimeInMs & " ms")
        If partItemType.isError Then
            Assert.Fail("Failed to connect")
        End If

    End Sub


    <TestMethod>
    Public Sub InsertHistory()
        Dim isCreated As Boolean = False
        Dim ITEM_NUMBER As String = PART_ITEM_NUMBER
        Dim message As String = "Testing history X"

RetryS:
        Dim timer As New Hille.Aras.Core.Utils.Timer.TaskTimer
        Dim amlQueryPart As String = "<AML><Item action='get' type='Part'><item_number>{0}</item_number></Item></AML>"
        amlQueryPart = String.Format(amlQueryPart, ITEM_NUMBER)

        Dim part As Item = Inn.applyAML(amlQueryPart)
        Console.WriteLine("Retrieved part in: " & timer.Restart & " ms")

        If part.isError And Not isCreated Then
            ' Try Create part
            Dim newPart As Item = Inn.applyAML(String.Format("<AML><Item action='add' type='Part'><item_number>{0}</item_number><name>{1}</name></Item></AML>", ITEM_NUMBER, PART_NAME))
            isCreated = True
            Console.WriteLine("Created part in: " & timer.Restart & " ms")
            part = newPart
            'GoTo RetryS
        End If

        If (Not part.isError()) Then
            timer.Restart()
            Dim history As History = New History(part)
            history.AddHistory(message)
            Console.WriteLine("Added history in: " & timer.GetElapsedTimeInMs & " ms")
            Console.WriteLine("Added history")
            If Not PrintHistory(history, message) Then
                Assert.Fail("Could not find message")
            End If
            ' Find message

        Else
            Console.WriteLine(part.getErrorString)
            Assert.Fail("Could not find part.")
        End If

    End Sub

    Private Function PrintHistory(history As History, checkForMessage As String) As Boolean
        Dim foundMessage As Boolean = False
        Dim list As List(Of HistoryEvent) = history.GetFullHistory
        For Each entry As HistoryEvent In list
            Dim outputString As String = entry.Action & ": " & entry.Comment & " (" & entry.WhoName & ") "
            Console.WriteLine(outputString)
            If entry.Comment = checkForMessage Then
                foundMessage = True
            End If
        Next
        Return foundMessage
    End Function
End Class
