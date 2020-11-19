Public Class PPAAudit
	Private Sub PPAAudit_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "PPA Audit"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTCDCDC"
			.vsql = "SET NOCOUNT ON
SELECT 
'PPA Daily Mix Match PDC Terms Audit' [Report Title],  
[master].[number] AS [File Number], 
cast([MiscExtraPPATerms].[thedata] as money) AS [PPA Terms], 
convert(varchar,[master].[lastpaid],101) AS [Last Payment Date], 
convert(varchar,[master].[lastpaidamt],101) AS [Last Payment Amount], 
CONVERT(varchar,nextpdc.deposit,101) [Next Due Date],
CONVERT(varchar,nextpdc.amount,101) [Next Due Amount]
FROM [master] WITH (NOLOCK)
INNER JOIN MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)
ON ([master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and isnumeric([MiscExtraPPATerms].[thedata]) = 1)
INNER JOIN [desk] WITH (NOLOCK)
ON ([master].[desk] = [desk].[code])
INNER JOIN [status] WITH (NOLOCK)
ON ([master].[status] = [status].[code])
join pdc_view_fbcs nextpdc with(nolock) on nextpdc.number=master.number
and nextpdc.UID=(select top 1 p.uid 
	from pdc_view_fbcs p with(nolock) 
	where p.number=master.number and p.Active=1
	and p.deposit>@start
	order by p.deposit )
WHERE ([desk].[branch] IN ('00002')
AND [master].[desk] IN ('100A', '101A', '102A', '103A', '104A', '105A', '106A', '107A', '10G', '19U', '20D', '20H', '36E', 'PPASPECIAL')
AND LEFT([status].[statustype], 1) = '0'
AND NOT(not exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%') OR exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and (thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0))))
AND ([master].[lastpaid] >= DATEADD(day,-1,@start )
and 
CAST(
case 
when NOT(thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
<> nextpdc.amount
)
union SELECT 
'PPA Daily Mix Match PDCC Terms Audit' [Report Title],  
[master].[number] AS [File Number], 
cast([MiscExtraPPATerms].[thedata] as money) AS [PPA Terms], 
convert(varchar,[master].[lastpaid],101) AS [Last Payment Date], 
convert(varchar,[master].[lastpaidamt],101) AS [Last Payment Amount], 
CONVERT(varchar,nextpdcc.depositdate,101) [Next Due Date],
CONVERT(varchar,nextpdcc.amount,101) [Next Due Amount]
FROM [master] WITH (NOLOCK)
INNER JOIN MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)
ON ([master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and isnumeric([MiscExtraPPATerms].[thedata]) = 1)
INNER JOIN [desk] WITH (NOLOCK)
ON ([master].[desk] = [desk].[code])
INNER JOIN [status] WITH (NOLOCK)
ON ([master].[status] = [status].[code])
join pdcc_view_fbcs nextpdcc with(nolock) on nextpdcc.number=master.number
and nextpdcc.ID=(select top 1 p.ID 
	from pdcc_view_fbcs p with(nolock) 
	where p.number=master.number and p.isActive=1
	and p.depositdate>@start
	order by p.depositdate )
WHERE ([desk].[branch] IN ('00002')
AND [master].[desk] IN ('100A', '101A', '102A', '103A', '104A', '105A', '106A', '107A', '10G', '19U', '20D', '20H', '36E', 'PPASPECIAL')
AND LEFT([status].[statustype], 1) = '0'
AND NOT(not exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%') OR exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and (thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0))))
AND ([master].[lastpaid] >= DATEADD(day,-1,@start )
and 
CAST(
case 
when NOT(thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
<> nextpdcc.amount
)
union SELECT 
'PPA Daily Mix Match Promise Terms Audit' [Report Title],  
[master].[number] AS [File Number], 
cast([MiscExtraPPATerms].[thedata] as money) AS [PPA Terms], 
convert(varchar,[master].[lastpaid],101) AS [Last Payment Date], 
convert(varchar,[master].[lastpaidamt],101) AS [Last Payment Amount], 
CONVERT(varchar,nextprom.duedate,101) [Next Due Date],
CONVERT(varchar,nextprom.amount,101) [Next Due Amount]
FROM [master] WITH (NOLOCK)
INNER JOIN MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)
ON ([master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and isnumeric([MiscExtraPPATerms].[thedata]) = 1)
INNER JOIN [desk] WITH (NOLOCK)
ON ([master].[desk] = [desk].[code])
INNER JOIN [status] WITH (NOLOCK)
ON ([master].[status] = [status].[code])
join promises_view_fbcs nextprom with(nolock) on nextprom.acctid=master.number
and nextprom.ID=(select top 1 p.ID 
	from promises_view_fbcs p with(nolock) 
	where p.acctid=master.number and p.Active=1
	and p.duedate>@start
	order by p.duedate )
WHERE ([desk].[branch] IN ('00002')
AND [master].[desk] IN ('100A', '101A', '102A', '103A', '104A', '105A', '106A', '107A', '10G', '19U', '20D', '20H', '36E', 'PPASPECIAL')
AND LEFT([status].[statustype], 1) = '0'
AND NOT(not exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%') OR exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and (thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0))))
AND ([master].[lastpaid] >= DATEADD(day,-1,@start )
and 
CAST(
case 
when NOT(thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
<> nextprom.amount
)
union SELECT 
'PPA Daily Mix Match Last Payment Terms Audit' [Report Title],  
[master].[number] AS [File Number], 
cast([MiscExtraPPATerms].[thedata] as money) AS [PPA Terms], 
convert(varchar,[master].[lastpaid],101) AS [Last Payment Date], 
convert(varchar,[master].[lastpaidamt],101) AS [Last Payment Amount],
'' [Next Due Date],
'' [Next Due Amount]
FROM [master] WITH (NOLOCK)
INNER JOIN MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)
ON ([master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and isnumeric([MiscExtraPPATerms].[thedata]) = 1)
INNER JOIN [desk] WITH (NOLOCK)
ON ([master].[desk] = [desk].[code])
INNER JOIN [status] WITH (NOLOCK)
ON ([master].[status] = [status].[code])
WHERE (
[desk].[branch] IN ('00002')
AND [master].[desk] IN ('100A', '101A', '102A', '103A', '104A', '105A', '106A', '107A', '10G', '19U', '20D', '20H', '36E', 'PPASPECIAL')
AND LEFT([status].[statustype], 1) = '0'
AND NOT(not exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%') OR exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and (thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)))
)
AND ([master].[lastpaid] >= DATEADD(day,-1,@start ))
and not exists(select * from PDC_View_FBCS p with(nolock) where p.number=master.number and p.Active=1
)
and not exists(select * from PDCC_View_FBCS p with(nolock) where p.number=master.number and p.isActive=1
)
and not exists(select * from Promises_View_FBCS p with(nolock) where p.acctid=master.number and p.Active=1
)
and not (
CAST(
case 
when NOT(thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
-
master.lastpaidamt between -50.00 and 50.00)
"
		End With
	End Sub
End Class