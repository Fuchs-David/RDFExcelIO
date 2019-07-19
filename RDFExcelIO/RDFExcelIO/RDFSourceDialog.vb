Public Class RDFSourceDialog
    Private Shared DEFAULT_LIMIT As Integer = 100
    Private sourceType As Boolean
    Private SPARQL_Address As Uri
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        sourceType = True
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Enabled = True
        TextBox2.Visible = False
        Button1.Enabled = False
        Label1.Visible = False
        Label2.Visible = False
        CheckBox1.Checked = False
        CheckBox1.Visible = False
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        sourceType = False
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox1.Enabled = True
        TextBox2.Visible = True
        Button1.Enabled = True
        Label1.Visible = True
        Label2.Visible = True
        CheckBox1.Visible = True
    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        If sourceType Then
            OpenFileDialog.ShowDialog()
        Else
            Label1.Visible = False
        End If
    End Sub

    Private Sub TextBox1_Leave(sender As Object, e As EventArgs) Handles TextBox1.Leave
        If Not sourceType And TextBox1.Text.Equals(String.Empty) Then
            Label1.Visible = True
        End If
    End Sub

    Private Sub TextBox2_GotFocus(sender As Object, e As EventArgs) Handles TextBox2.GotFocus
        Label2.Visible = False
    End Sub

    Private Sub TextBox2_Leave(sender As Object, e As EventArgs) Handles TextBox2.Leave
        If TextBox2.Text.Equals(String.Empty) Then
            Label2.Visible = True
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If sourceType Then
            If RDFExcelIO.SpecifyDataSource(TextBox1.Text) Then
                DialogResult = Windows.Forms.DialogResult.Yes
            End If
        Else
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            Try
                Dim limit As Integer
                Dim query As String = ""
                Dim currentOption As String = ""
                If Not Integer.TryParse(TextBox2.Text, limit) Then
                    limit = DEFAULT_LIMIT
                End If
                If RadioButton1.Checked Then
                    currentOption = "file"
                ElseIf RadioButton2.Checked Then
                    currentOption = "SPARQL"
                End If
                If CheckBox1.Checked Then
                    query = TextBox3.Text
                    currentOption = "customSPARQL"
                End If
                If RDFExcelIO.SpecifyDataSource(New Uri(TextBox1.Text), limit, query, currentOption) Then
                    DialogResult = Windows.Forms.DialogResult.Yes
                End If
            Catch ex As Exception
                MsgBox("Invalid SPARQL endpoint URL", MsgBoxStyle.Critical)
            End Try
        End If
        Close()
    End Sub

    Private Sub OpenFileDialog_FileOk(sender As Object, e As ComponentModel.CancelEventArgs) Handles OpenFileDialog.FileOk
        TextBox1.Text = OpenFileDialog.FileName
        TextBox1.Enabled = False
        Button1.Enabled = True
    End Sub

    Private Sub RDFSourceDialog_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButton1.Text = LocalizeText("fileInput")
        RadioButton2.Text = LocalizeText("sparql")
        Label1.Text = LocalizeText("sparql")
        Label2.Text = LocalizeText("limit")
        Button1.Text = LocalizeText("confirmSource")
        CheckBox1.Text = LocalizeText("useCustomQuery")
        Select Case RDFExcelIO.GetLastOption
            Case "file"
                RadioButton1.Checked = True
            Case "SPARQL"
                RadioButton2.Checked = True
                Label1.Visible = False
            Case "customSPARQL"
                RadioButton2.Checked = True
                CheckBox1.Checked = True
                Label1.Visible = False
        End Select
        OpenFileDialog.Filter = "RDF|*.rdf|TriG|*.trig|TriX|*.trix|NTriples|*.nt|Turtle|*.ttl|All files|*.*"
        TextBox1.Text = ToSafeString(RDFExcelIO.GetLastEndpoint)
        TextBox3.Text = RDFExcelIO.GetLastQuery
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If TextBox3.ReadOnly Then
            TextBox3.ReadOnly = False
            TextBox2.Clear()
            TextBox2.ReadOnly = True
        Else
            TextBox3.ReadOnly = True
            TextBox2.ReadOnly = False
        End If
    End Sub
End Class