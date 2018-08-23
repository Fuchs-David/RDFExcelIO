Imports VDS.RDF
Public Class SelectSubjects
    Private newSheetIndicator As Boolean = False
    Private Sub SelectMainDimension_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CheckBox1.Text = LocalizeText("newSheet")
        Button1.Text = LocalizeText("useSelectedItems")
        ListBox1.SelectionMode = Windows.Forms.SelectionMode.MultiExtended
        For Each node As INode In RDFExcelIO.GetNodes
            ListBox1.Items.Add(node)
        Next
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        newSheetIndicator = Not newSheetIndicator
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim selection As New Queue(Of INode)
        For Each item In ListBox1.SelectedItems
            selection.Enqueue(CType(item, INode))
        Next
        RDFExcelIO.ConvertRDF(selection, newSheetIndicator)
        Close()
    End Sub
End Class