Imports VDS.RDF
' Container class for mapping nodes to lines
Public Class INodeToRow
    Private node As INode
    Private row As Integer
    Public Sub New(ByRef node As INode, ByVal row As Integer)
        Me.node = node
        Me.row = row
    End Sub
    Public Function GetNode() As INode
        Return node
    End Function
    Public Function GetRow() As Integer
        Return row
    End Function
End Class
