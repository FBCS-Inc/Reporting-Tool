Public Class BrokenPromises
	Private Sub BrokenPromises_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.RPTNAME = "Settlements to be 2nd reviewed"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTTTDCCTCCC"
			.Refresh()
			.Show()

		End With
	End Sub
End Class