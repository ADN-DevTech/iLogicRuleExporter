Imports System.IO

Public Class InventorInfo

	Private inventorApp As Inventor.Application
	Public m_quitInventor As Boolean = False

	Public Sub StartInventor()
		Try
			' Try to get an active instance of Inventor
			Try
				Debug.Print("StartInventor - Trying to find Inventor application")
				inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
				Debug.Print(" - Found it")
			Catch ex As Exception
			End Try

			' Don't create a new Inventor session. We need an active document.
			If inventorApp Is Nothing Then
				MessageBox.Show("You must start Inventor and open a document before exporting rules", "iLogic Rule Exporter")
				'LaunchInventor()
			End If

		Catch ex As Exception
            MessageBox.Show("There was a problem starting Inventor", "Error", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
        End Try
	End Sub

	Public Sub StopInventor()
		Debug.Print("StopInventor")
		If m_quitInventor = True Then
			Debug.Print(" - Quitting Inventor")
			inventorApp.Quit()
			WaitForInventorToDie()
		End If

		inventorApp = Nothing
	End Sub

	Friend ReadOnly Property Application() As Inventor.Application
		Get
			Return inventorApp
		End Get
	End Property


	Sub LaunchInventor()
		Debug.Print("InventorInfo - Launching Inventor...")
		Dim inventorAppType As Type = System.Type.GetTypeFromProgID("Inventor.Application")
		inventorApp = System.Activator.CreateInstance(inventorAppType)
		'm_inventorAppEvents = inventorApp.ApplicationEvents

		WaitForInventorToBeReady()

		inventorApp.Visible = True
		m_quitInventor = True
	End Sub

	Sub WaitForInventorToBeReady()
		Dim inventorReady As Boolean = False
		While Not inventorReady
			System.Threading.Thread.Sleep(100)
			System.Windows.Forms.Application.DoEvents()
			If inventorApp.Ready Then inventorReady = True
		End While
		Debug.Print("InventorInfo - Inventor Ready")
	End Sub

	Sub WaitForInventorToDie()
		' Wait for Inventor to be gone
		Do While Not inventorApp Is Nothing
			Try
				inventorApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application")
				System.Threading.Thread.Sleep(100)
			Catch ex As Exception
				Debug.Print(" - Inventor is dead")
				inventorApp = Nothing
			End Try
		Loop
	End Sub

End Class

