Public Module GetiLogicAddin

  Function GetiLogicAutomation(ByVal app As Inventor.Application) As Object

        Dim addIns As Inventor.ApplicationAddIns = app.ApplicationAddIns

        Try
            Dim customAddIn As Inventor.ApplicationAddIn = addIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
            Return customAddIn.Automation

        Catch ex As Exception
            MessageBox.Show("Could not find the iLogic addin", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return Nothing

    End Function

End Module
