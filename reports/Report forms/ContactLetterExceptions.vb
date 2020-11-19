
Public Class ContactLetterExceptions
	Private Sub ContactLetterExceptions_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.RPTNAME = "Contact Letter Exceptions"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + ".xlsx"
			.shfmt = "TT"

			.vsql = "declare @report table(
number integer not null,
contacted datetime not null,
reason varchar(50) not null)

insert into @report(number,contacted,reason)
select master.number,isnull(master.contacted,master.worked),
case 
when NOT( ([Debtors].[MR] = 'N' OR [debtors].[MR] = '0')
	AND NOT [Debtors].[Street1] = ''
	AND NOT [Debtors].[Street1] IS NULL
	AND NOT [Debtors].[City] = ''
	AND NOT [Debtors].[City] IS NULL
	AND NOT [Debtors].[State] = ''
	AND NOT [Debtors].[State] IS NULL
	AND NOT [Debtors].[ZipCode] = ''
 	AND NOT [Debtors].[ZipCode] IS NULL) 
 then 'Bad Address'
 when exists(select * from notes with(nolock) where notes.number=master.number and notes.created>=master.contacted and notes.comment like '%Mail Return Set%')
 then 'Bad Address'
 when EXISTS(SELECT * FROM restrictions WITH (NOLOCK) WHERE restrictions.number=master.number AND restrictions.suppressletters = 1)
 then 'Letter Restrictions'
 when exists(select * from notes with(nolock) where notes.number=master.number and notes.created>=master.contacted and notes.result in ('FB','DE','FD','D','PD','VD'))
 then 'FB,DE,FD, D, PD or VD Result used'
when CustomCustGroups.Name like 'ops/%Center%Point%'
AND CASE WHEN ISNUMERIC([miscextra].[thedata]) = 1 THEN CAST([miscextra].[thedata] AS INTEGER) ELSE 0 END < 600 then 'Center Point Score Under 600'
 when exists(select * from notes with(nolock) where notes.number=master.number and notes.comment like '%Account reinstated%')
 then 'Account Reinstated'
 when master.desk  in (select code from restricteddesks with(nolock)) and master.desk not in ('101A','103A','105A','19U','36E')
 then 'Restricted Desk ' + master.desk
 when desk.branch in (select branch from excludedbranches with(nolock)) and desk.Branch not in (12,2)
 then 'Restricted Branch ' + desk.branch
when exists(Select * from payhistory with(nolock) where payhistory.number=master.number and payhistory.batchtype in ('PU','PC'))
then 'Account has made a payment'
when
link>0 and link is not null and 
(CASE WHEN [master].[link] = 0 OR [master].[link] = '' OR ([master].[link] IS NULL) THEN 1 ELSE (select count(mlink.number) from master mLink with (nolock) join fact flink with(nolock) on flink.customerid=mlink.customer join status statusLink with (nolock) on (statusLink.code = mLink.status) where mLink.link = [master].[link] and fact.customgroupid=flink.customgroupid and left(statusLink.statustype,1) = 0) END between 2 and 9 OR CASE WHEN [master].[link] = 0 OR [master].[link] = '' OR ([master].[link] IS NULL) THEN 1 ELSE (select count(mlink.number) from master mLink with (nolock) join fact flink with(nolock) on flink.customerid=mlink.customer join status statusLink with (nolock) on (statusLink.code = mLink.status) where mLink.link = [master].[link] and fact.customgroupid=flink.customgroupid and left(statusLink.statustype,1) = 0) END IS NULL)
	and (
		customcustgroups.Name like 'ops/Pallino%'
		or customcustgroups.Name like 'ops/Pendrick%'
		or customcustgroups.Name like 'ops/Capio%'
		or customcustgroups.Name like 'ops/Cascade%'
		or customcustgroups.Name like 'ops/FFAM%'
	)
then 'Received ' + convert(varchar(15),master.received,101) +   '    Last Link Letter ' + isnull((select top 1 convert(varchar(15),l.daterequested,101) from letterrequest l with(nolock) join master m with(nolock) on m.number=l.accountid join fact f with(nolock) on f.customerid=m.customer
where f.customgroupid=fact.customgroupid and m.link=master.link and l.deleted=0 and l.lettercode in (60,61,62,63,66) and l.daterequested>master.received order by l.daterequested desc),'')
else ''
 end
from master with(nolock)
join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
join desk with(nolock) on desk.code=master.desk
join fact with(nolock) on fact.customerid=master.customer
join customcustgroups with(nolock) on customcustgroups.id=fact.customgroupid and customcustgroups.name like 'ops/%'
LEFT OUTER JOIN [miscextra] WITH (NOLOCK)
	ON ([miscextra].[number] = [master].[number] and ([miscextra].[title] = 'TU_MiscInfo1' or [miscextra].[title] = 'TUMiscInfo1'))
where  (
contacted between DATEADD(day,-1,{fn curdate()}) and DATEADD(ss,-1,{fn curdate()})
/*or (
	worked between DATEADD(day,-1,{fn curdate()}) and DATEADD(ss,-1,{fn curdate()})
	and exists(select * from notes with(nolock) where notes.number=master.number
	and notes.created between DATEADD(day,-1,{fn curdate()}) and DATEADD(ss,-1,{fn curdate()})
	and notes.result='HU')
   )*/
)
and closed is null
and not exists(select * from letterrequest lr with(Nolock)
join letter l with(nolock) on l.code=lr.LetterCode
where lr.AccountID=master.number
and (l.Description like '%letter of rep%' or l.Description like 'LOR %' or l.Description like '% LOR%')
and lr.deleted=0
and lr.suspend=0
)

select number,reason
from @report
order by reason"
		End With

	End Sub
End Class