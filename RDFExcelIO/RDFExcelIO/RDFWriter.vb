Imports VDS.RDF
Imports VDS.RDF.Writing
Imports Microsoft.Office.Interop.Excel
Public Class RDFWriter
    Implements IDisposable
    Private Const firstRow As Integer = 1
    Private Const firstCol As Integer = 1

    Private currentSheet As Worksheet
    Private currentSheetCells As Range
    Private graph As IGraph
    Private rdfxmlWriter As New RdfXmlWriter

    Public Function CreateGraph(ByRef baseURI As Uri) As Boolean
        graph = New Graph()
        graph.BaseUri = baseURI
        graph.NamespaceMap.Clear()
        currentSheet = CType(Globals.RDFExcelIO.Application.ActiveWorkbook.ActiveSheet, Worksheet)
        currentSheetCells = currentSheet.Cells
        CreateTriples(firstRow, firstCol, firstRow, firstCol)
        If graph.IsEmpty Then
            graph.Dispose()
            Return False
        End If
        Return True
    End Function
    Private Sub CreateTriples(ByVal firstRow As Integer, ByVal firstCol As Integer, ByVal row As Integer, ByVal col As Integer)
        Dim r As Integer = row + 1
        While Not CType(currentSheetCells(r, col), Range).Value2 Is Nothing
            If Uri.IsWellFormedUriString(CType(currentSheetCells(r, col), Range).Value2.ToString, UriKind.Absolute) Then
                Dim subject As IUriNode = graph.CreateUriNode(New Uri(CType(currentSheetCells(r, col), Range).Value2.ToString))
                Dim c As Integer = col + 1
                While Not CType(currentSheetCells(r, c), Range).Value2 Is Nothing
                    If firstCol <> c AndAlso CType(currentSheetCells(row, c), Range).Value2.Equals("Subject") Then
                        CreateTriples(firstRow, col, row, c)
                    ElseIf Uri.IsWellFormedUriString(CType(currentSheetCells(firstRow, c), Range).Value2.ToString, UriKind.Absolute) _
                       AndAlso
                       Uri.IsWellFormedUriString(CType(currentSheetCells(r, c), Range).Value2.ToString, UriKind.Absolute) Then
                        graph.Assert(New Triple(subject,
                              graph.CreateUriNode(New Uri(CType(currentSheetCells(firstRow, c), Range).Value2.ToString)),
                              graph.CreateUriNode(New Uri(CType(currentSheetCells(r, c), Range).Value2.ToString))))
                    ElseIf Uri.IsWellFormedUriString(CType(currentSheetCells(firstRow, c), Range).Value2.ToString, UriKind.Absolute) _
                           AndAlso
                           CType(currentSheetCells(r, c), Range).Hyperlinks.Count = 1 _
                           AndAlso
                           Uri.IsWellFormedUriString(CType(currentSheetCells(r, c), Range).Hyperlinks.Item(1).Address, UriKind.Absolute) Then
                        Dim objectNode As ILiteralNode = graph.CreateLiteralNode(CType(currentSheetCells(r, c), Range).Value2.ToString,
                                                                                 New Uri(CType(currentSheetCells(r, c), Range).Hyperlinks.Item(1).Address & "#" &
                                                                                 CType(currentSheetCells(r, c), Range).Hyperlinks.Item(1).SubAddress))
                        graph.Assert(New Triple(subject, graph.CreateUriNode(New Uri(CType(currentSheetCells(firstRow, c), Range).Value2.ToString)),
                                                objectNode))
                    ElseIf Uri.IsWellFormedUriString(CType(currentSheetCells(firstRow, c), Range).Value2.ToString, UriKind.Absolute) Then
                        Dim objectNode As ILiteralNode = graph.CreateLiteralNode(CType(currentSheetCells(r, c), Range).Value2.ToString)
                        graph.Assert(New Triple(subject, graph.CreateUriNode(New Uri(CType(currentSheetCells(firstRow, c), Range).Value2.ToString)),
                                                objectNode))
                    End If
                    c += 1
                End While
            End If
            r += 1
        End While
    End Sub
    Public Sub WriteGraphToFile(ByRef fileName As String)
        rdfxmlWriter.Save(graph, fileName)
        graph.Dispose()
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                graph.Dispose()
            End If

            currentSheet = Nothing
            currentSheetCells = Nothing
        End If
        disposedValue = True
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
    End Sub
#End Region
End Class
