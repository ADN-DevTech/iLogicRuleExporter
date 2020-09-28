Public Class Form1
    Private Sub exportRulesButton_Click(sender As Object, e As EventArgs) Handles exportRulesButton.Click

        Using (New WaitCursor())
            Dim exp As New RulesExporter()
            exp.ExportAllRules()
            exp.ReleaseInventor()
        End Using

    End Sub
End Class
