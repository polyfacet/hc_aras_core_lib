Namespace BomCompare.OutputFormat

    Public Class DefaultOutputColors
        Implements IOutputColors

        Public Sub New()
            Me.AddedRowColor = "lightgreen"
            Me.RemovedRowColor = "pink"
            Me.ReplacedRowColor = "lightorange"
            ' Individiual cell
            Me.ChangedCellColor = "yellow"
            Me.SplitterCellColor = "#000000"
        End Sub

        Public Property AddedRowColor As String Implements IOutputColors.AddedRowColor
        Public Property ChangedCellColor As String Implements IOutputColors.ChangedCellColor
        Public Property ChangedRowColor As String Implements IOutputColors.ChangedRowColor
        Public Property RemovedRowColor As String Implements IOutputColors.RemovedRowColor
        Public Property ReplacedRowColor As String Implements IOutputColors.ReplacedRowColor
        Public Property SplitterCellColor As String Implements IOutputColors.SplitterCellColor

    End Class

End Namespace