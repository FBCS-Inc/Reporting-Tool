Public Class Removed737Q
	Private Sub Removed737Q_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "Removed From 737 Queue"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTTTTT"
			.vsql = "select debtors.Number, debtors.Email, 
case when len(isnull(debtors.SSN,''))< 4 then '' else RIGHT(debtors.ssn,4) end [SSN], 
desk.code [Desk Code], desk.name [Desk Name], isnull(rejectReason.thedata,'') [Other Reason] 
from MiscExtra with(nolock) 
join master with(nolock) on master.number=miscextra.number 
join debtors with(nolock) on debtors.Number=miscextra.Number and debtors.Seq=0 
left outer join miscextra rejectReason with(nolock) on rejectReason.number=miscextra.number and rejectReason.title='EmailAddError' 
join desk with(nolock) on desk.code=master.desk 
where master.closed is null and miscextra.Title='EmailAddErrorDate' 
and case when ISDATE(miscextra.thedata)=1 then CAST(miscextra.thedata as datetime) else '1-1-1753' end >= @start) "
		End With
	End Sub
End Class