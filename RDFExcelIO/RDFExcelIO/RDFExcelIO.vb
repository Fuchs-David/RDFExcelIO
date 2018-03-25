Imports System.Windows.Forms
Imports VDS.RDF
Public Class RDFExcelIO
    Private Shared reader As New RDFReader
    Private Shared writer As New RDFWriter
    Private Shared rdfSourcgDialog As RDFSourceDialog
    Private Shared subjectDialog As SelectSubjects
    Private Shared saveFileDialog As New SaveFileDialog

    Public Shared Sub ConvertRDF(ByRef node As Queue(Of INode), ByVal newSheetIndicator As Boolean)
        reader.ConvertRDFToTable(node, newSheetIndicator)
    End Sub
    Public Shared Sub ConvertToRDF()
        If saveFileDialog.ShowDialog = DialogResult.OK Then
            Dim baseURI As Uri = reader.GetGraphURI
            reader.ClearData()
            If writer.CreateGraph(baseURI) Then
                writer.WriteGraphToFile(saveFileDialog.FileName)
            End If
        End If
    End Sub
    Public Shared Function GetNodes() As HashSet(Of INode)
        Return reader.GetSubjectNodes
    End Function
    Public Shared Function SpecifyDataSource(SPARQLAddress As Uri, limit As Integer) As Boolean
        Return reader.ConnectToSPARQLEndpoint(SPARQLAddress, limit)
    End Function
    Public Shared Function SpecifyDataSource(fileName As String) As Boolean
        Return reader.ReadFile(fileName)
    End Function
    Public Shared Function GetRDFDataSource() As Boolean
        rdfSourcgDialog = New RDFSourceDialog
        Return rdfSourcgDialog.ShowDialog() = DialogResult.Yes
    End Function
    Public Shared Sub GetMainDataDimension()
        subjectDialog = New SelectSubjects
        subjectDialog.Show()
    End Sub

    Protected Overrides Function CreateRibbonExtensibilityObject() _
                                    As Microsoft.Office.Core.IRibbonExtensibility
        saveFileDialog.AddExtension = True
        saveFileDialog.FileName = "file"
        saveFileDialog.Filter = "RDF file|*.rdf"
        saveFileDialog.DefaultExt = "rdf"
        Return New Ribbon()
    End Function
End Class
