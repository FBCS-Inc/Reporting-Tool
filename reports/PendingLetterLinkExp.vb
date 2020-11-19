

Public Class PendingLetterLinkExp
   
    Private Sub PendingLetterLinkExp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .StartLabel.Visible = False
            .StartDate.Visible = False
            .enddate.Text = Now.Date.ToString
            .EndLabel.Text = "Through Date"
            .RPTNAME = " Pending Letters - Link Exceptions"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTTDT"
            .vsql = "declare @report table(
number integer not null,
lettercode varchar(10) not null,
daterequested datetime not null,
comment varchar(250) not null
)
declare @ForceDisasterLetters table(
letterrequestid integer not null,
number integer not null,
debtorid integer not null,
allow bit not null
)
insert into @ForceDisasterLetters(letterrequestid,number,debtorid,allow)
select distinct lr.letterrequestid, m.number, d.DebtorID, 1
from letterrequest lr with (nolock)
	inner join letterrequestrecipient lrr with (nolock) on lr.letterrequestid = lrr.letterrequestid
	inner join letter l with (nolock) on l.letterid = lr.letterid
	inner join master m with (nolock) on m.number = lr.accountid
	left outer join debtors d with (nolock) on d.debtorid = lrr.debtorid
	join restrictions r with (nolock) on r.debtorid = lrr.debtorid
	left outer join letterdisclaimermapping lmap with(nolock) on lmap.[Letter Code]=lr.lettercode
	WHERE lr.DateRequested <= @end  
	AND (lr.DateProcessed IS NULL OR lr.DateProcessed = '1/1/1753 12:00:00')
	AND lr.Deleted = 0 AND lr.AddEditMode = 0 AND lr.Suspend = 0 AND lr.Edited = 0 
	--AND lr.LetterID = @LetterID
	and lr.LetterCode like '31%'
	and r.suppressletters=1
	and lmap.[disasterapproved]=1
	and (select top 1 notes.comment from notes with(nolock) 
	where notes.number=lr.AccountID
	--and notes.comment like '%Debtor%' + CAST(d.seq+1 as varchar) + '%'
	and (
	notes.comment like '%Debtor%' + CAST(d.seq+1 as varchar) + '%'
	or notes.comment like '%due to storm%'
	)
	and (
		notes.comment like '%Restriction%Add%due%to%'
		or notes.comment like '%Restriction%Add%due%to%release%'
		or notes.comment like '%restrict letters set%'
		or notes.comment like '%restrict letters cleared%'
	)
	order by notes.created desc
	)
 like '%Restriction%Add%due%to%'

declare @linkErrors table(
number int not null,
debtorid int not null,
link int not null,
daterequested datetime not null,
lettercode varchar(10) not null,
errordescription varchar(30) not null)

insert into @linkErrors(number,debtorid,link,daterequested,lettercode,errordescription)
select sd.number, sd.debtorid, m.link,lr.daterequested,lr.lettercode,
		CASE	
			--10/10/18 per jamie barton force reminder letters through	
			WHEN (r.suppressletters = 1) and ( forceletter.allow is null or forceletter.allow = 0 ) THEN 'Letter Suppressed'
			WHEN (l.allowclosed = 0) and (m.qlevel = '998' or m.qlevel = '999') THEN 'Account closed'
			WHEN (l.allowzero = 0) and (m.current0 = 0) THEN 'Account at zero balance'
			WHEN (m.current0 < 0) THEN 'Account at negative balance'
			WHEN (RTRIM(ltrim(d.name))=',' or RTRIM(ltrim(d.name))='' or d.Name is null) THEN 'Account has no name'
			WHEN (l.allowfirst30 = 1) and (l.allowafter30 = 0) and (getdate() > dateadd(d, 30, m.received)) THEN 'Cannot send letter after the first 30 days'
			WHEN (l.allowafter30 = 1) and (l.allowfirst30 = 0) and (getdate() < dateadd(d, 30, m.received)) THEN 'Cannot send letter during the first 30 days'
			WHEN ((sr.restricted = 1) and (l.type <> 'ATT' or l.type <> 'CUS')) THEN 'Letters restricted in ' + sr.statename
			WHEN l.emailsubject = 'Approved' and sd.ssn is not null	and sd.ssn<>'' and LEN(sd.SSN) >= 4
			and emailsubscription.email is not null	and emailsubscription.email <> '' and emailsubscription.suspended is null
			and (r.letterstoatty is null or r.letterstoatty = 0) and (lrr.AltRecipient is null or lrr.AltRecipient = 0)
			THEN 'Email'
			WHEN (l.type <> 'ATT' or l.type <> 'CUS') and (d.mr = 'Y') THEN 'Bad Address'
			WHEN (l.type <> 'ATT' or l.type <> 'CUS') and (len(rtrim(ltrim(d.street1))) < 5) THEN 'Bad Address'
			ELSE ''
		END
	from letterrequest lr with (nolock)
	inner join letterrequestrecipient lrr with (nolock) on lr.letterrequestid = lrr.letterrequestid
	inner join letter l with (nolock) on l.letterid = lr.letterid
	inner join master m with (nolock) on m.number = lr.accountid
	left outer join debtors d with (nolock) on d.debtorid = lrr.debtorid
	inner join customer cu with (nolock) on cu.customer = m.customer
	left outer join restrictions r with (nolock) on r.debtorid = lrr.debtorid
	left outer join attorney a with (nolock) on a.attorneyid = m.attorneyid
	left outer join courtcases cc with (nolock) on cc.accountid = lr.accountid
	left outer join courts co with (nolock) on co.courtid = cc.courtid
	left outer join StateRestrictions sr with (nolock) on sr.abbreviation = d.State
	inner join desk de with (nolock) on de.code = m.desk
	inner join branchcodes bc with (nolock) on bc.code = de.branch
	left outer join debtors sd with (nolock) on sd.debtorid = lr.subjdebtorid
	--added fbcs 1/5/15
	left outer join emailsubscription with(nolock) on emailsubscription.debtorid=sd.debtorid
	left outer join @ForceDisasterLetters forceLetter on forceLetter.letterrequestid = lr.LetterRequestID and forceLetter.debtorid=d.DebtorID
	WHERE lr.DateRequested <= @end  
	AND (lr.DateProcessed IS NULL OR lr.DateProcessed = '1/1/1753 12:00:00')
	AND lr.Deleted = 0 AND lr.AddEditMode = 0 
	AND lr.Suspend = 0 AND lr.Edited = 0 
	and lr.LetterCode like '31%'
insert into @report(number,lettercode,daterequested,comment)	
select distinct l1.number,l1.lettercode,l1.daterequested,
'Link Error ' +
(select top 1 l2.errordescription from @linkErrors l2 where l1.link=l2.link and l1.number<>l2.number and l1.debtorid<>l2.debtorid
and l2.errordescription<>'')
from @linkErrors l1 
where l1.errordescription=''
and exists(select * from @linkErrors l2 where l1.link=l2.link and l1.number<>l2.number and l1.debtorid<>l2.debtorid
and l2.errordescription<>'')
insert into @report(number,lettercode,daterequested,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),
'Cover ' + convert(varchar,AmountDue,101) + ' not matching Details ' + CONVERT(varchar, isnull((select SUM(l2.amountdue) from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and dbo.fngetday(l2.dateprocessed)=dbo.fngetday(letterrequest.dateprocessed)
and l2.LetterCode='31004'),0.00),101) 
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.accountid
where LetterCode in (31001, 31003, 31007)
and DateRequested<=@end 
and YEAR(dateprocessed)=1753
and Deleted=0
and 
NOT (AmountDue - isnull((select SUM(l2.amountdue) from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link 
and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and l2.LetterCode='31004'),0.00)
between -1 and 1)
and not exists(select * from @report r where r.number=accountid)
insert into @report(number,lettercode,daterequested,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),
'Cover ' + CONVERT(varchar,
case When(isnumeric(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))))=1) then cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))))=1) then cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))))=1) then cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))))=1) then cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))))=1) then cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))))=1) then cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))))=1) then cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))))=1) then cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))))=1) then cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))))=1) then cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))))=1) then cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))))=1) then cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))))=1) then cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))))=1) then cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))))=1) then cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))))=1) then cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))))=1) then cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))))=1) then cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))))=1) then cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))))=1) then cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))))=1) then cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))))=1) then cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))))=1) then cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))))=1) then cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money) else 0.00 end 
,101) + ' not matching Details ' +
CONVERT(varchar,
isnull(
(select 
SUM(
case When(isnumeric(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))))=1) then cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))))=1) then cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))))=1) then cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))))=1) then cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))))=1) then cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))))=1) then cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))))=1) then cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))))=1) then cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))))=1) then cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))))=1) then cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))))=1) then cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))))=1) then cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))))=1) then cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))))=1) then cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))))=1) then cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))))=1) then cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))))=1) then cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))))=1) then cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))))=1) then cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))))=1) then cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))))=1) then cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))))=1) then cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))))=1) then cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))))=1) then cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money) else 0.00 end 
)
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link 
and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and l2.LetterCode='31005'),0.00),101)
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.accountid
where LetterCode in (31000, 31002, 31006)
and DateRequested<=@end 
and YEAR(dateprocessed)=1753
and Deleted=0
--the total on the cover does not match the sum of the details
and 
NOT(
case When(isnumeric(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))))=1) then cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))))=1) then cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))))=1) then cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))))=1) then cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))))=1) then cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))))=1) then cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))))=1) then cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))))=1) then cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))))=1) then cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))))=1) then cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))))=1) then cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))))=1) then cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))))=1) then cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))))=1) then cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))))=1) then cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))))=1) then cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))))=1) then cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))))=1) then cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))))=1) then cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))))=1) then cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))))=1) then cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))))=1) then cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))))=1) then cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))))=1) then cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money) else 0.00 end 
-
 isnull((select 
SUM(
case When(isnumeric(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))))=1) then cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))))=1) then cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))))=1) then cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))))=1) then cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))))=1) then cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))))=1) then cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))))=1) then cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))))=1) then cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))))=1) then cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))))=1) then cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))))=1) then cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))))=1) then cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))))=1) then cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))))=1) then cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))))=1) then cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))))=1) then cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))))=1) then cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))))=1) then cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))))=1) then cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))))=1) then cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))))=1) then cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))))=1) then cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))))=1) then cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money) else 0.00 end +
case When(isnumeric(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))))=1) then cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money) else 0.00 end 
)
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link 
and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and l2.LetterCode='31005'),0.00)
between -1 and 1 )
and not exists(select * from @report r where r.number=accountid)

insert into @report(number,lettercode,daterequested,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),
case when
master.MR in ('1','Y')
or RTRIM(master.street1)=''
or master.Street1 is null
or RTRIM(master.city)=''
or master.city is null
or RTRIM(master.state)=''
or master.state is null
or RTRIM(master.Zipcode)=''
or master.Zipcode is null then 'Bad Address On This Link'
when exists(select * from restrictions with(nolock) where restrictions.number=master.number and restrictions.suppressletters=1) then
'Letter Restrictions On This Link'
else ''
end 
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.AccountID
join debtors with(nolock) on debtors.number=master.number and debtors.debtorid=letterrequest.subjdebtorid
where letterrequest.LetterCode like '31%'
and DateRequested<=@end 
and YEAR(dateprocessed)=1753
and Deleted=0
and (
master.MR in ('1','Y')
or RTRIM(master.street1)=''
or master.Street1 is null
or RTRIM(master.city)=''
or master.city is null
or RTRIM(master.state)=''
or master.state is null
or RTRIM(master.Zipcode)=''
or master.Zipcode is null
or exists(select * from restrictions with(nolock) where restrictions.number=master.number and debtors.debtorid=restrictions.debtorid and restrictions.suppressletters=1)
)
and exists(select * 
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.AccountID
join debtors d2 with(nolock) on d2.number=m2.number and d2.debtorid=l2.subjdebtorid
where l2.LetterCode like '31%'
and dbo.fnGetDay(l2.daterequested)=dbo.fnGetDay(letterrequest.daterequested)
and l2.deleted=0
and m2.link=master.link
and m2.number<>master.number
and NOT(
m2.MR in ('1','Y')
or RTRIM(m2.street1)=''
or m2.Street1 is null
or RTRIM(m2.city)=''
or m2.city is null
or RTRIM(m2.state)=''
or m2.state is null
or RTRIM(m2.Zipcode)=''
or m2.Zipcode is null
or exists(select * from restrictions with(nolock) where restrictions.number=m2.number
and restrictions.debtorid=d2.debtorid
 and restrictions.suppressletters=1)
)
)
and not exists(select * from @report r where r.number=accountid)

select distinct master.link,r.*
from @report r 
join master with(nolock) on master.number=r.number
order by master.link,r.number"
        End With

    End Sub

End Class