Public Class ExeterBankruptcyCloses
	Private Sub ExeterBankruptcyCloses_Load(sender As Object, e As EventArgs) Handles MyBase.Load
			With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "Exeter Bankruptcy Closes"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "T"
			.vsql = "select distinct master.number as [File Number] from master 
INNER join notes (NOLOCK) on notes.number = master.number
left outer join bankruptcy (NOLOCK) on bankruptcy.AccountID = master.number
left outer join debtorattorneys (NOLOCK) on debtorattorneys.AccountID = master.number
INNER JOIN fact with(NOLOCK) ON fact.customerid=master.customer 
INNER JOIN customcustgroups c with(NOLOCK) ON c.id=fact.customgroupid  and c.name like 'FIN/Exeter'
where(
notes.created > @start and notes.result = 'FB'
and
((bankruptcy.CaseNumber is NULL or bankruptcy.CaseNumber = '') OR (debtorattorneys.Name is NULL or debtorattorneys.Name = '' or debtorattorneys.Name like '%refuse%')) 
)"
		End With
	End Sub
End Class