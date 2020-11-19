
Public Class Exception7
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With ReportData1
            .RPTNAME = "Operations Exceptions 7 - Contacts"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTCTTTTDD"
            .vsql = "select top 2500 master.number [File #], opsgroup.name [Client Ops Group],
master.current0 [Balance],master.State [State],
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
case when master.contacted is null then '' else convert(varchar,master.contacted,101) end [Last Contact],
(Select top 1 convert(varchar,priorCt.created,101) from notes priorCt with(nolock)
	join result r2 with(nolock) on priorCt.result=r2.code
	where priorCt.number=master.number
	and priorCt.action not in ('DT','LOG')
	and dbo.fnGetDay(priorCt.created)<dbo.fnGetDay(master.contacted)
	and dbo.fnGetDay(priorCt.created)<>dbo.fnGetDay(master.contacted)
	and r2.contacted=1
	and DATEDIFF(day,priorCt.created,master.contacted)<=8
	order by dbo.fngetday(priorCt.created) desc
	) [Prior Contact]
from master with(nolock)
join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
join customer with(nolock) on customer.customer=master.customer
join desk with(nolock) on desk.code=master.desk
INNER JOIN [miscextra] score WITH (NOLOCK)
ON (score.[number] = [master].[number] and (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
join Fact with(nolock) on fact.CustomerID=master.customer
join CustomCustGroups opsgroup with(nolock) on opsgroup.ID=fact.CustomGroupID and opsgroup.Name like 'ops%'
where 
contacted is not null
--Account was contacted in the last 7 days (weekly report)
and contacted>=DATEADD(day,-7,@start)
--Open Status
and closed is null
--Account was contacted within 8 days of the last contact
and exists(select *
	from notes c1 with(nolock)
	join result r1 with(nolock) on c1.result=r1.code
	where c1.number=master.number
	and c1.action not in ('DT','LOG')
	and r1.contacted=1
	and dbo.fnGetDay(master.contacted)=dbo.fnGetDay(c1.created)
	and exists(Select * from notes c2 with(nolock)
	join result r2 with(nolock) on c2.result=r2.code
	where c2.number=master.number
	and c2.action not in ('DT','LOG')
	and dbo.fnGetDay(c2.created)<dbo.fnGetDay(c1.created)
	and dbo.fnGetDay(c2.created)<>dbo.fnGetDay(c1.created)
	and c2.UID<>c1.uid
	and r2.contacted=1
	and DATEDIFF(day,c2.created,c1.created)<=8
	)
)"
        End With

    End Sub


End Class