Public Class Autobuildercount
	Private Sub Autobuildercount_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "LiveVox AutoBuilder Phone Counts"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTDTTTTTTTTTTTT"
			.vsql = "select substring(replace(queryFileName,'_HCI',''),0,30) [Query File], sum(AccountCount ) [Accounts], 
dbo.fngetday(dateBuilt) dateBuilt , sum(homephoneCount) [Home Phones], sum(workphoneCount) [Work Phones], 
sum(altphone1Count) [Cell Phones], sum(altphone2Count) [Skip Phone 1], sum(altphone3Count) [Skip Phone 2],
sum(altphone4Count) [Skip Phone 3], sum(altphone5Count) [Alt Cell phone], sum(altphone6Count) [Alt Home Phone 1], 
sum(altphone7Count) [Alt Home Phone 2], sum(altphone8Count) [Alt Home Phone 3], sum(altphone9Count ) [Alt Home Phone 4],
sum(altphone10Count) [Alt Home Phone 5] 
from LiveVoxAutoBuilderHistory with(nolock) 
where datebuilt >= @start
group by substring(replace(queryFileName,'_HCI',''),0,30), dbo.fngetday(datebuilt) 
order by left(substring(replace(queryFileName,'_HCI',''),0,30),2),cast(dbo.fnStripNonDigits(substring(replace(queryFileName,'_HCI',''),0,30)) as Int)
"
		End With
	End Sub
End Class