Public Class Exception4
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Operations Exceptions 4 - First 30 Days"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
            .shfmt = "TTCTTTTDTDD"
            .vsql = "declare @workDays int
set @workDays=5
declare @contactDays int
set @contactDays=8

declare @accounts table(
number int not null,
link int null,
lastworked datetime null
)

insert into @accounts(number,link,lastworked)
select master.number,master.link,(
case when link=0 or link is null then 
isnull(master.worked,'1900-01-01') 
else 
isnull((select top 1 m2.worked from master m2 with(nolock) 
	where m2.linkdriver in (0,1) 
	--and m2.qlevel<'998' 
	and m2.link=master.link order by m2.worked desc),'1900-01-01') 
end) 
from master with(nolock)
join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
join customer with(nolock) on customer.customer=master.customer
join desk with(nolock) on desk.code=master.desk
INNER JOIN [miscextra] score WITH (NOLOCK)
ON (score.[number] = [master].[number] and (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
join Fact with(nolock) on fact.CustomerID=master.customer
join CustomCustGroups opsgroup with(nolock) on opsgroup.ID=fact.CustomGroupID and opsgroup.Name like 'ops%'
where  
--Placed in the last 30 days
received >=   DATEADD(day,-30,@start) 
--its been at least 5 business days since placement
and (
DATEDIFF(d,master.received,@start) 
- DATEDIFF(wk,master.received,@start) * 2 
- CASE WHEN DATENAME(dw, master.received) <> 'Saturday' AND DATENAME(dw, @start) = 'Saturday' THEN 1 
WHEN DATENAME(dw, master.received) = 'Saturday' AND DATENAME(dw, @start) <> 'Saturday' THEN -1 
ELSE 0 
END ) 
>= 5
--Branch is not 2
and desk.Branch<>'00002'
--Open Status
and closed is null
/* 12-02-2019 added by e. hall per K. Buono to exclude ops/JCAP Auto Value Tier D and ops/JCAP Auto Value Tier E */
and  opsgroup.name not in ('ops/JCAP Auto Value Tier D','ops/JCAP Auto Value Tier E')
--No restrictions
and (not exists(select * from bankruptcy with(nolock) where bankruptcy.accountid=master.number and bankruptcy.datefiled <> '1753-01-01 12:00:00.000' and bankruptcy.datefiled<>'1900-01-01 00:00:00.000') and not exists(select * from deceased with(nolock) where deceased.accountid=master.number and deceased.dod is not null and deceased.dod <> '1753-01-01 12:00:00.000' and deceased.dod <> '1900-01-01 00:00:00.000') and not exists (select * from debtorattorneys with(nolock) where debtorattorneys.accountid=master.number and ( rtrim(debtorattorneys.[name]) <> '' or rtrim(debtorattorneys.[firm]) <> '' or rtrim(debtorattorneys.Addr1) <> '' or rtrim(debtorattorneys.Addr2) <> '' or rtrim(debtorattorneys.City) <> '' or rtrim(debtorattorneys.[State]) <> '' or rtrim(debtorattorneys.Zipcode) <> '' or rtrim(debtorattorneys.Phone) <> '' or rtrim(debtorattorneys.Fax) <> '' or rtrim(debtorattorneys.Email) <> '' or rtrim(debtorattorneys.Comments) <> '') ) and not exists(select * from restrictions with(nolock) where restrictions.number=master.number and ((not(home is null) and home <> 0) OR (not(job is null) and job <> 0)  OR  (not(calls is null) and calls <> 0) OR (not(comment is null) and not(comment like '')) OR (not(AttyName is null) and AttyName <> '') OR (not(Attyaddr1 is null) and Attyaddr1 <> '') OR (not(Attyaddr2 is null) and Attyaddr2 <> '') OR (not(AttyCity is null) and AttyCity <> '') OR (not(AttyState is null) and AttyState <> '') OR (not(AttyZip is null) and AttyZip <> '') OR (not(AttyPhone is null) and AttyPhone <> '') OR (not(Attynotes is null) and not(Attynotes like ''))OR (not(BkyCase is null) and BkyCase <> '') OR (not(BkyChap11 is null) and BkyChap11 <> 0) OR (not(BkyChap13 is null) and BkyChap13 <> 0) OR (not(BkyChap7 is null) and BkyChap7 <> 0) OR (not(BkyCourt is null) and BkyCourt <> '') OR (not(BkyDateFiled is null)) OR (not(BkyDistrict is null) and BkyDistrict <> '') OR (not(BkyNotes is null) and not(BkyNotes like '')) OR (not(suppressletters is null) and suppressletters <> 0) OR (not(Disputed is null) and Disputed <> 0))))
--Not in a restricted desk
and master.desk not in (Select rd.code from restricteddesks rd with(nolock))
--Has a phone in good or unknown status thats not a poe
and exists(select* from Phones_Master with(nolock) 
where phones_master.Number=master.number 
and phones_master.debtorid=debtors.debtorid and phones_master.OnHold=0
and (phones_master.PhoneStatusID is null or phones_master.PhoneStatusID=2)
and phones_master.phonetypeid not in (2)
)
--Last worked date is more than 5 business days ago
and (
master.worked is null
or 
(
(DATEDIFF(d,master.worked,@start) --Get the number of days between start and end dates
- DATEDIFF(wk,master.worked,@start) * 2 -- for each week, subtract 2 days (by default a week occurs between sat and sunday on sql server)
- CASE WHEN DATENAME(dw, master.worked) <> 'Saturday' AND DATENAME(dw, @start) = 'Saturday' THEN 1 --subtract 1 day if the end date falls on a saturday and the startdate is a weekday, or sunday --written by pflangan haha
WHEN DATENAME(dw, master.worked) = 'Saturday' AND DATENAME(dw, @start) <> 'Saturday' THEN -1 --add 1 if the start date is a saturday and the end date is any other day
ELSE 0 
END ) > @workDays
)
)
--Last contact date is more than 8 days ago
and (
master.contacted is null
or 
(
DATEDIFF(d,master.contacted,@start) > @contactDays
)
)
--Score 600+
and CASE WHEN ISNUMERIC(score.[thedata]) = 1 THEN CAST(score.[thedata] AS INTEGER) ELSE 0 END >= 600
--No Active Arrangements
and not exists(select * from PDC_View_FBCS pdc with(nolock) where pdc.number=master.number and pdc.Active=1)
and not exists(select * from PDCC_View_FBCS pdc with(nolock) where pdc.number=master.number and pdc.isActive=1)
and not exists(select * from promises with(nolock) where promises.acctid=master.number and promises.Active=1)
and not exists(select * from paymentbatchitems p with(nolock) where p.FileNum=master.number and p.PmtType=1)
--Last Good Posititve Payment was more than 30 days ago
and isnull((select top 1 entered from payhistory p1 with(nolock) 
	where p1.number=master.number and p1.batchtype in ('PU','PC') 
	and not exists(select * from payhistory p1rev with(nolock) where p1rev.number=master.number
	and p1rev.batchtype in ('PUR','PCR')
	and p1rev.ReverseOfUID=p1.UID)
	order by entered desc),'1-1-1900')
< DATEADD(day,-30,@start)
--Exclude desks Spanish, Spanish14, A0115, PENDCLSE & All Collector Desks
and master.desk not in ('Spanish', 'Spanish14', 'A0115', 'PENDCLOSE')
and desk.desktype not in ('Collector')
--Also exclude clients MCM & Turnstile
and opsgroup.name not like '%ops%midland%'
and opsgroup.name not like '%ops%turnstile%'

select top 2500 master.number [File #], opsgroup.name [Client Ops Group],
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
convert(varchar,master.Received,101) [Received],master.Status,
isnull(convert(varchar,master.worked,101),'') [Last Worked], 
case when a.lastworked='1-1-1900' then '' else convert(varchar,a.lastworked,101) end [Link Last Worked] 
from @accounts a
join master with(nolock) on master.number=a.number
join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
join customer with(nolock) on customer.customer=master.customer
join desk with(nolock) on desk.code=master.desk
INNER JOIN [miscextra] score WITH (NOLOCK)
ON (score.[number] = [master].[number] and (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
join Fact with(nolock) on fact.CustomerID=master.customer
join CustomCustGroups opsgroup with(nolock) on opsgroup.ID=fact.CustomGroupID and opsgroup.Name like 'ops%'
where  (
a.lastworked is null
or 
(
(DATEDIFF(d,a.lastworked,@start) --Get the number of days between start and end dates
- DATEDIFF(wk,a.lastworked,@start) * 2 -- for each week, subtract 2 days (by default a week occurs between sat and sunday on sql server)
- CASE WHEN DATENAME(dw, a.lastworked) <> 'Saturday' AND DATENAME(dw, @start) = 'Saturday' THEN 1 --subtract 1 day if the end date falls on a saturday and the startdate is a weekday, or sunday --written by pflangan haha
WHEN DATENAME(dw, a.lastworked) = 'Saturday' AND DATENAME(dw, @start) <> 'Saturday' THEN -1 --add 1 if the start date is a saturday and the end date is any other day
ELSE 0 
END ) > @workDays
)
)
order by master.received desc"
        End With
    End Sub

End Class