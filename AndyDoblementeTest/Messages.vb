Public Class Messages
    Public Shared Sub ShowErrorMessageBox(text As String)
        Dim message As String = "INFORMATION"
        Dim buttons As MessageBoxButtons = MessageBoxButtons.OK
        Dim icon As MessageBoxIcon = MessageBoxIcon.Error
        Dim result As DialogResult = MessageBox.Show(text, message, buttons, icon)
    End Sub
End Class
