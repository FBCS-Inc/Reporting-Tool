Public Class Arrang45days
	Private Sub Arrang45days_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "Arrangements greater than 45 days"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTDDC"
			.vsql = "select master.number [File Number],
case 
when (DATEDIFF(day,@start,pdc.deposit)>=45 and pdc.deposit is not null)
then 'PDC'
when (DATEDIFF(day,@start,pdcc.depositdate)>=45 and pdcc.depositdate is not null)
then 'PDCC'
when (DATEDIFF(day,@start,promises.duedate)>=45 and promises.duedate is not null)
then 'Promise'
else '' end [Arrangement Type],
case 
when (DATEDIFF(day,@start,pdc.deposit)>=45 and pdc.deposit is not null)
then convert(varchar,pdc.entered,101)
when (DATEDIFF(day,@start,pdcc.depositdate)>=45 and pdcc.depositdate is not null)
then convert(varchar,pdcc.dateentered,101)
when (DATEDIFF(day,@start,promises.duedate)>=45 and promises.duedate is not null)
then convert(varchar,promises.entered,101)
else '' end [Entered Date],
case 
when (DATEDIFF(day,@start,pdc.deposit)>=45 and pdc.deposit is not null)
then convert(varchar,pdc.deposit,101)
when (DATEDIFF(day,@start,pdcc.depositdate)>=45 and pdcc.depositdate is not null)
then convert(varchar,pdcc.depositdate,101)
when (DATEDIFF(day,@start,promises.duedate)>=45 and promises.duedate is not null)
then convert(varchar,promises.duedate,101)
else '' end [Next Due Date],
case 
when (DATEDIFF(day,@start,pdc.deposit)>=45 and pdc.deposit is not null)
then convert(varchar,pdc.amount,101)
when (DATEDIFF(day,@start,pdcc.depositdate)>=45 and pdcc.depositdate is not null)
then convert(varchar,pdcc.amount,101)
when (DATEDIFF(day,@start,promises.duedate)>=45 and promises.duedate is not null)
then convert(varchar,promises.amount,101)
else '' end [Next Due Amount]
from master with(nolock)
left outer join PDC_View_FBCS pdc with(nolock) on pdc.UID=
	(select top 1 p.uid from PDC_View_FBCS p with(nolock)
		where master.number=p.number and p.Active=1
		order by p.deposit)
left outer join PDCC_View_FBCS pdcc with(nolock) on pdcc.ID=
	(select top 1 p.id from PDCC_View_FBCS p with(nolock)
		where master.number=p.number and p.isActive=1
		order by p.depositdate)
left outer join Promises_View_FBCS promises with(nolock) on promises.ID=
	(select top 1 p.id from promises_View_FBCS p with(nolock)
		where master.number=p.acctid and p.Active=1
		order by p.duedate)				
where closed is null
and (pdc.UID is not null or pdcc.ID is not null or promises.ID is not null)
and 
(DATEDIFF(day,@start,pdc.deposit)>=45 or pdc.deposit is null)
and
(DATEDIFF(day,@start,pdcc.depositdate)>=45 or pdcc.depositdate is  null)
and
(DATEDIFF(day,@start,promises.duedate)>=45 or promises.duedate is  null)
 "

		End With
	End Sub
End Class