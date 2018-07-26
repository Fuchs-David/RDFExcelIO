Module Utility
    Public Function ToSafeString(ByVal objectToString As Object) As String
        If objectToString Is Nothing Then
            Return ""
        Else
            Return objectToString.ToString
        End If
    End Function
End Module
