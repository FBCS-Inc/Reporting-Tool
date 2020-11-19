Public Class PCDDeletedReport
    Private Sub PCDDeletedReport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Post Dated Transactions Deleted On Closed Accounts"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTDDDTTTDT"
            .StartDate.Visible = False
            .StartLabel.Visible = False
            .enddate.Visible = False
            .EndLabel.Visible = False
            .vsql = "declare @StartDate datetime
declare @EndDate datetime
set @StartDate = dateadd(day,-1,{fn curdate()})
set @EndDate = dateadd(ss,-1,{fn curdate()})
declare @report table(
number integer not null,
status varchar(10) not null,
closed datetime  null,
returned datetime  null,
created datetime not null,
comment varchar(max) not null,
desk varchar(10) not null,
ppacreditdesk varchar(10) not null,
arrangDate datetime null,
customer varchar(100) not null)
insert into @report(number,status,closed,returned,created,comment,desk,ppacreditdesk,arrangDate,customer)
select notes.number,master.status, master.closed, master.returned,notes.created,notes.comment,master.desk [CurrentDesk],
CASE 
when desk.branch<>'00002' and desk.code<>'105A' then master.desk
WHEN (desk.code='105A' AND NOT(ppacredit.thedata is null or ppacredit.thedata = master.desk)) 
AND (CASE WHEN NOT([ppacreditdate].[theData] IS NULL) AND RTRIM(LTRIM([ppacreditdate].[theData])) <> '' 
          THEN 
               CASE WHEN isdate([ppacreditdate].[theData])=1 and ([ppacreditdate].[theData] LIKE '%-%-%' OR [ppacreditdate].[theData] LIKE '%/%/%' ) THEN [ppacreditdate].[theData] ELSE NULL END
     ELSE NULL END) >=  [desk_to].[Hire_Date]
and (Select top 1 p1.batchtype 
	from payhistory p1 with(nolock) where p1.number=master.number 
	and p1.batchtype in ('PU','PC','PUR','PCR') Order by p1.entered desc, p1.uid desc) in ('PU','PC')
then ppacredit.thedata
WHEN (desk.branch='00002' AND NOT(ppacredit.thedata is null or ppacredit.thedata = master.desk)) 
AND (CASE WHEN NOT([ppacreditdate].[theData] IS NULL) AND RTRIM(LTRIM([ppacreditdate].[theData])) <> '' 
          THEN 
               CASE WHEN isdate([ppacreditdate].[theData])=1 and ([ppacreditdate].[theData] LIKE '%-%-%' OR [ppacreditdate].[theData] LIKE '%/%/%' ) THEN [ppacreditdate].[theData] ELSE NULL END
     ELSE NULL END) >=  [desk_to].[Hire_Date]
then ppacredit.thedata
else master.desk end [ppacreditdesk],
case when notes.comment like '%date of %' then
right(
left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2),
LEN(left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2))
-
CHARINDEX('date of',left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2),0)
-7
)
when notes.comment like '%due on %' then
right(
left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2),
LEN(left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2))
-
CHARINDEX('due on',left(CAST(notes.comment as varchar(5000)),CHARINDEX('12:00AM',CAST(notes.comment as varchar(5000)),0)-2),0)
-6
) End,
customer.name
from master with(nolock)
join customer with(nolock) on customer.customer=master.customer
join notes with(nolock) on notes.number=master.number and notes.user0='Custodian' 
INNER JOIN [desk] WITH (NOLOCK) ON master.[desk] = [desk].[code]
LEFT OUTER JOIN miscextra ppacredit with(nolock) on ppacredit.number=master.number and ppacredit.title like 'PPA%Credit%Desk' and exists(select * from desk d with(nolock) where d.code=rtrim(ppacredit.thedata))
LEFT OUTER JOIN [MiscExtra] AS [ppacreditdate] WITH (NOLOCK) ON master.[number] = [ppacreditdate].[number]AND [ppacreditdate].[Title] = 'PPACreditDate'
LEFT OUTER JOIN [desk_to] with (nolock) ON ([desk_to].[desk] = [ppacredit].[TheData])
where 
/*closed between @StartDate and dateadd(hh,23,@EndDate) removed 02-19-2018*/
master.status not in ('SIF','PIF')
and notes.comment like '%deleted because account was closed%'
and notes.created>=@StartDate 
order by notes.created,notes.number


select *
from @report
where arrangDate>closed
or arrangDate is null
order by number,arrangDate"
        End With
    End Sub
End Class