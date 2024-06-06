Imports System.Text.RegularExpressions

Public Class StringValidations
    Public Shared Function HasNoWhitespace(value As String)
        Dim regex As New Regex("^\S+$")
        Return regex.IsMatch(value)
    End Function
End Class
