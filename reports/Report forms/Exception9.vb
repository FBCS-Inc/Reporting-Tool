
Public Class Exception9

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Operations Exceptions 9 - Dialer Contacted in less than 8 days"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTCTTTTDD"
            .vsql = "select lastct.accountid,lastCt.CallDateTime cdt 
into #lastContact
from livevoxresults lastCt with(nolock) 
inner join [result] with(nolock) on [result].code = lastct.Result and [result].contacted=1
where lastCt.[Connected]=1
and lastCt.Result <> 'DH'
and lastCt.CallDateTime>=DATEADD(day,-7,@start)
and lastCt.CallDateTime = (select max(L.CallDateTime) from livevoxresults l with(nolock)inner join [result] with(nolock) on [result].code = l.Result and [result].contacted=1 where l.accountid = lastCt.accountid)

select #lastContact.accountid,#lastContact.cdt [LC],priorCt.CallDateTime [PC]
into #contactn8
from #lastContact with(nolock) 
inner join livevoxresults priorCt with(nolock) on #lastContact.accountid = priorCt.accountid and DATEDIFF(dayofyear,priorCt.CallDateTime,#lastContact.cdt)<8
inner join [result] with(nolock) on [result].code = priorCt.Result and [result].contacted=1
where #lastContact.cdt > priorCt.CallDateTime
and  priorCt.Result <> 'DH'


select top 2500 master.number [File #], opsgroup.name [Client Ops Group],master.current0 [Balance],master.State [State],
score.thedata [Score],master.Desk + ': ' + desk.name [Desk],
case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2) and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0 and phones_master.PhoneTypeID=1) then 'Home ' else '' end
+
case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2) and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0 and phones_master.PhoneTypeID=2) then 'Work ' else '' end
+
case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2) and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0 and phones_master.PhoneTypeID=3) then 'Cell ' else '' end
+
case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2) and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0 and phones_master.PhoneTypeID=7) then 'Skip ' else '' end
+
case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2) and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0 and phones_master.PhoneTypeID not in (1,2,3,7)) then 'Other ' else '' end
[Phone Types],
convert(varchar,#contactn8.LC,101) [Last Contact], convert(varchar,#contactn8.[PC],101) [Prior Contact]
from master with(nolock)
inner join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
inner join customer with(nolock) on customer.customer=master.customer
inner join desk with(nolock) on desk.code=master.desk
INNER JOIN [miscextra] score WITH (NOLOCK)ON (score.[number] = [master].[number] and (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
inner join Fact with(nolock) on fact.CustomerID=master.customer
inner join CustomCustGroups opsgroup with(nolock) on opsgroup.ID=fact.CustomGroupID and opsgroup.Name like 'ops%'
inner join #contactn8 with(nolock) on #contactn8.accountid = master.number
where master.closed is null
Drop Table #lastContact
Drop Table #contactn8"
        End With

    End Sub

End Class