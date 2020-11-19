Public Class recordings2weeks
	Private Sub recordings2weeks_Load(sender As Object, e As EventArgs) Handles MyBase.Load
				With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "LiveVox Recordings 2 Week Summary"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "DTT"
			.vsql = "select dbo.fnGetDay(callstarttime)[Call Start],datename(dw,callstarttime)[Weekday],COUNT(*) [Recordings]
from LiveVoxRecordingDetails l with(nolock) 
where callstarttime between DATEADD(day,-14,@start) and @start 
group by dbo.fnGetDay(callstarttime),datename(dw,callstarttime) order by dbo.fnGetDay(callstarttime)"
		End With
	End Sub
End Class