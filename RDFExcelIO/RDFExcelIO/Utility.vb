Imports System.Text.RegularExpressions

Module Utility
    Private cultures As String = "en-US|cs-CZ"
    Public Function ToSafeString(ByVal objectToString As Object) As String
        If objectToString Is Nothing Then
            Return ""
        Else
            Return objectToString.ToString
        End If
    End Function
    Public Function LocalizeText(ByVal identifier As String) As String
        Dim result As String = "en_us"
        If Regex.IsMatch(Threading.Thread.CurrentThread.CurrentUICulture.Name, cultures) Then
            result = Threading.Thread.CurrentThread.CurrentUICulture.Name.ToLower.Replace("-", "_")
        End If
        Return My.Resources.ResourceManager.GetString(result & "_" & identifier)
    End Function
End Module
