Public Class Recordingcount
	Private Sub Recordingcount_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "LiveVox Recordings Counts"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "DTT"
			.vsql = "select dbo.fngetday(callstarttime),count(*) 
from LiveVoxRecordingDetails with(nolock) 
where LiveVoxRecordingDetails.callstarttime>=@start group by dbo.fngetday(callstarttime)"
		End With
	End Sub
End Class