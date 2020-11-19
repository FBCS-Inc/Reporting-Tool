Public Class PPAException
	Private Sub PPAException_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.StartDate.Visible = False
			.StartLabel.Visible = False
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "PPA Exception"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TCDCCT"
			.vsql = "SELECT [master].[number] AS [File Number], 
cast([MiscExtraPPATerms].[thedata] as money) AS [PPA Terms], 
convert(varchar,[master].[lastpaid],101) AS [Last Payment Date], 
convert(varchar,[master].[lastpaidamt],101) AS [Last Payment Amount],
convert(varchar,[master].current0,101) AS [Current Balance],
master.Desk [Current Desk]
FROM [master] WITH (NOLOCK)
INNER JOIN MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)ON ([master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and isnumeric([MiscExtraPPATerms].[thedata]) = 1)
INNER JOIN [desk] WITH (NOLOCK)ON ([master].[desk] = [desk].[code])
INNER JOIN [status] WITH (NOLOCK)ON ([master].[status] = [status].[code])
WHERE ([desk].[branch] IN ('00002')
--AND [master].[desk] IN ('100A', '101A', '102A', '103A', '104A', '105A', '106A', '107A', '10G', '19U', '20D', '20H', '36E', 'PPASPECIAL')
AND LEFT([status].[statustype], 1) = '0'
AND NOT(not exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%') OR exists(select * from MiscExtra AS [MiscExtraPPATerms] WITH (NOLOCK)where  [master].[number]=[MiscExtraPPATerms].[number] and [MiscExtraPPATerms].[title] like '%PPA%TERMS%' and (thedata = '' or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)))
)
and CAST(
case 
when NOT(thedata =''  or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
> current0
and NOT(
CAST(
case 
when NOT(thedata =''  or thedata = '0' or isnumeric([MiscExtraPPATerms].[thedata]) = 0)
then [MiscExtraPPATerms].[thedata]
else '0' end as Money)
- current0
) between -1 and 1
and master.desk<>'PPAPENDCLS'
and thedata <>''  
and thedata <> '0'
and isnumeric(thedata)=1"
		End With
	End Sub
End Class