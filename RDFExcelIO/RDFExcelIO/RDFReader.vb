Imports VDS.RDF
Imports VDS.RDF.Parsing
Imports VDS.RDF.Query
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing
Imports System.Web
Imports System.Text.RegularExpressions

Public Class RDFReader
    Implements IDisposable
    Private Shared QP_LIMIT As String = "limit"
    Private Shared QP_SUBJECT As String = "subject"
    Private Shared exploratoryQuery As String = "describe ?subject
WHERE  { ?subject ?predicate ?object } LIMIT @" & QP_LIMIT
    Private Shared detailQuery As String = "describe @" & QP_SUBJECT & "
WHERE  { @" & QP_SUBJECT & " ?predicate ?object } LIMIT @" & QP_LIMIT

    Private Const firstRow As Integer = 1
    Private Const firstCol As Integer = 1

    Private graph As IGraph
    Private SPARQLEndpoint As SparqlRemoteEndpoint = Nothing
    Private currentSheet As Worksheet
    Private currentSheetCells As Range
    Private limit As Integer ' Limit of SPARQL results
    Private predicatesLeadingToRecursion As HashSet(Of INode)
    Public Function GetGraphURI() As Uri
        Return graph.BaseUri
    End Function
    Public Sub ClearData()
        graph.Dispose()
        SPARQLEndpoint = Nothing
    End Sub
    ' Read contents of RDF file.
    Public Function ReadFile(fileName As String) As Boolean
        graph = New Graph
        predicatesLeadingToRecursion = New HashSet(Of INode)
        FileLoader.Load(graph, fileName)
        If graph.IsEmpty Then
            graph.Dispose()
            Return False
        End If
        Return True
    End Function
    ' Get graph from SPARQL endpoint.
    Public Function ConnectToSPARQLEndpoint(endpoint As Uri, limit As Integer) As Boolean
        graph = New Graph
        predicatesLeadingToRecursion = New HashSet(Of INode)
        Me.limit = limit
        Try
            SPARQLEndpoint = New SparqlRemoteEndpoint(endpoint)
        Catch ex As Exception
            SPARQLEndpoint = Nothing
            graph.Dispose()
            MsgBox("Failed to connect to SPARQL endpoint.", MsgBoxStyle.Critical)
        End Try
        QuerySPARQLEndpoint(exploratoryQuery, QP_LIMIT, limit)
        Return Not graph.IsEmpty
    End Function
    Private Sub QuerySPARQLEndpoint(queryString As String, QP_name As String, QP_value As Integer)
        Try
            Dim query As SparqlParameterizedString = New SparqlParameterizedString(queryString)
            query.SetLiteral(QP_name, QP_value)
            graph.Merge(SPARQLEndpoint.QueryWithResultGraph(query.ToString), True)
        Catch ex As Exception
            MsgBox("Failed to query SPARQL endpoint.", MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub QuerySPARQLEndpoint(queryString As String, QP_name As String, QP_value As INode)
        Try
            Dim query As SparqlParameterizedString = New SparqlParameterizedString(queryString)
            query.SetParameter(QP_name, QP_value)
            query.SetLiteral(QP_LIMIT, limit)
            graph.Merge(SPARQLEndpoint.QueryWithResultGraph(query.ToString), True)
        Catch ex As HttpException
            MsgBox("Failed to query SPARQL endpoint.", MsgBoxStyle.Critical)
        End Try
    End Sub
    Public Function GetSubjectNodes() As HashSet(Of INode)
        Dim nodes As New HashSet(Of INode)
        For Each node As INode In graph.Triples.SubjectNodes
            nodes.Add(node)
        Next
        For Each node In graph.Triples.PredicateNodes
            nodes.Remove(node)
        Next
        Return nodes
    End Function
    Public Sub ConvertRDFToTable(ByVal nodes As Queue(Of INode), ByVal useNewSheet As Boolean)
        Dim row As Integer = firstRow + 1
        If useNewSheet Then
            currentSheet = CType(Globals.RDFExcelIO.Application.ActiveWorkbook.Worksheets.Add(),
                                 Worksheet)
        Else
            currentSheet = CType(Globals.RDFExcelIO.Application.ActiveWorkbook.ActiveSheet,
                                 Worksheet)
            currentSheet.UsedRange.Clear()
        End If
        currentSheetCells = currentSheet.Cells()
        While nodes.Any
            WriteRowToTable(nodes.Dequeue, row, firstCol)
            row += 1
        End While
    End Sub
    Private Sub WriteRowToTable(ByRef subjectNode As INode,
                                ByRef row As Integer, ByVal col As Integer)
        WriteTextToCell(firstRow, col, "Subject")
        WriteTextToCell(row, col, subjectNode.ToSafeString)
        col += 1
        Dim rowOriginal As Integer = row
        Dim RowMax As Integer = row
        For Each predicate In GetPredicateNodes(subjectNode)
            Dim c As Integer = col
            For Each triple In graph.Triples.WithSubjectPredicate(subjectNode, predicate)
                While Not CType(currentSheetCells(firstRow, c), Range).Value Is Nothing
                    If CType(currentSheetCells(firstRow, c), Range).Value.ToString _
                             .Equals(predicate.ToSafeString) Then
                        Exit While
                    Else
                        c += 1
                    End If
                End While
                If CType(currentSheetCells(firstRow, c), Range).Value Is Nothing Then
                    WriteNodeToCell(firstRow, c, predicate)
                End If
                WriteNodeToCell(row, c, triple.Object, graph.Triples.WithSubject(triple.Object).Count = 0,
                                        predicatesLeadingToRecursion.Contains(predicate))
                predicatesLeadingToRecursion.Add(predicate)
                AdjustRowMax(row, RowMax)
                row += 1
            Next
            row = rowOriginal
        Next
        row = RowMax
    End Sub
    Private Sub AdjustRowMax(ByVal row As Integer, ByRef rowMax As Integer)
        If row > rowMax Then
            rowMax = row
        End If
    End Sub
    Private Function GetPredicateNodes(ByRef node As INode) As HashSet(Of INode)
        Dim predicates As New HashSet(Of INode)
        For Each predicate In graph.Triples.PredicateNodes
            predicates.Add(predicate)
        Next
        Return predicates
    End Function
    Private Sub WriteTextToCell(ByVal row As Integer, ByVal col As Integer, ByRef text As String)
        currentSheetCells(row, col) = text
    End Sub
    Private Sub WriteLiteralNodeToCell(ByVal row As Integer, ByVal col As Integer,
                                       ByRef node As LiteralNode)
        If Regex.Match(node.DataType.ToSafeString, "double|float|hexBinary|decimal|" &
                    "integer|long|int|short|byte|" &
                    "nonNegativeInteger|positiveInteger|unsignedLong|unsignedInt|unsignedShort|unsignedByte|" &
                    "nonPositiveInteger|negativeInteger").Success Then
            CType(currentSheetCells(row, col), Range).Value2 = node.Value
        Else
            CType(currentSheetCells(row, col), Range).NumberFormat = "@"
            CType(currentSheetCells(row, col), Range).Value2 = node.Value
        End If
        AddHyperlink(CType(currentSheetCells(row, col), Range), node.DataType.ToSafeString)
    End Sub
    Private Sub WriteNodeToCell(ByRef row As Integer, ByVal c As Integer, ByRef node As INode,
                                Optional ByVal recurse As Boolean = False,
                                Optional ByVal doNotAskUser As Boolean = False)
        If node.NodeType = NodeType.Literal Then
            WriteLiteralNodeToCell(row, c, CType(node, LiteralNode))
        ElseIf node.NodeType = NodeType.Uri _
               AndAlso recurse _
               AndAlso (doNotAskUser OrElse MsgBox("Continue loading information about object nodes?",
                              MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
            If Not SPARQLEndpoint Is Nothing Then
                QuerySPARQLEndpoint(detailQuery, QP_SUBJECT, node)
            End If
            WriteTextToCell(row, c, node.ToSafeString)
            WriteRowToTable(node, row, c + 1)
        Else
            WriteTextToCell(row, c, node.ToSafeString)
        End If
    End Sub
    Private Sub AddHyperlink(ByRef cell As Range, ByRef address As String)
        If address.Equals(String.Empty) Then
            Exit Sub
        End If
        currentSheet.Hyperlinks.Add(cell, address)
        cell.Font.Color = Color.Black
        cell.Font.Underline = XlUnderlineStyle.xlUnderlineStyleNone
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                graph.Dispose()
            End If

            SPARQLEndpoint = Nothing
            predicatesLeadingToRecursion = Nothing
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
