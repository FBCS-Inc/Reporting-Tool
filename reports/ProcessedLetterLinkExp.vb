Public Class ProcessedLetterLinkExp

    Private Sub ProcessedLetterLinkExp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .StartLabel.Text = "Letter Date"
            .StartDate.Text = Now.Date.ToString
            .enddate.Visible = False
            .EndLabel.Visible = False
            .RPTNAME = " Processed Letters - Link Exceptions"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTTDT"
            .vsql = "
if @start >= '10-25-18' 
begin

declare @report table(
number integer not null,
lettercode varchar(10) not null,
daterequested datetime not null,
dateprocessed datetime  null,
comment varchar(250) not null
)

insert into @report(number,lettercode,daterequested,dateprocessed,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),dateprocessed,
letterrequest.ErrorDescription + ' Letter Console Exception not found on link'
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.AccountID
where letterrequest.LetterCode like '31%'
and isnull(letterrequest.ErrorDescription,'')<>''
and Dateprocessed between @start and dateadd(ss,-1,dateadd(d,1,@start)) -->={fn curdate()}--DATEADD(day,-1,{fn curdate()})
and Deleted=0
and exists(select * 
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.AccountID
where l2.LetterCode like '31%'
and dbo.fnGetDay(l2.daterequested)=dbo.fnGetDay(letterrequest.daterequested)
and m2.link=master.link
and m2.number<>master.number
and isnull(l2.ErrorDescription,'')=''
)
and not exists(select * from @report r where r.number=accountid)

insert into @report(number,lettercode,daterequested,dateprocessed,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),dateprocessed,
'Cover ' + convert(varchar,AmountDue,101) + ' not matching Details ' + CONVERT(varchar, isnull((select SUM(l2.amountdue) from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and dbo.fngetday(l2.dateprocessed)=dbo.fngetday(letterrequest.dateprocessed)
and l2.deleted=0
and l2.LetterCode='31004'),0.00),101) 
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.accountid
where LetterCode in (31001, 31003, 31007)
--and DateRequested>=DATEADD(day,-1,{fn curdate()})
and Dateprocessed between @start and dateadd(ss,-1,dateadd(d,1,@start)) -->={fn curdate()}--DATEADD(day,-1,{fn curdate()})
--and YEAR(dateprocessed)=1753
and Deleted=0
and NOT
(
AmountDue -
isnull((select SUM(l2.amountdue) from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and dbo.fngetday(l2.dateprocessed)=dbo.fngetday(letterrequest.dateprocessed)
and l2.LetterCode='31004'),0.00)
between -1 and 1)
and not exists(select * from @report r where r.number=accountid)



insert into @report(number,lettercode,daterequested,dateprocessed,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),dateprocessed,
'Cover ' + CONVERT(varchar,

'case When(isnumeric(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))))=1) then cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))))=1) then cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))))=1) then cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))))=1) then cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))))=1) then cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))))=1) then cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))))=1) then cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))))=1) then cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))))=1) then cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))))=1) then cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))))=1) then cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))))=1) then cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))))=1) then cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))))=1) then cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))))=1) then cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))))=1) then cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))))=1) then cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))))=1) then cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))))=1) then cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))))=1) then cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))))=1) then cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))))=1) then cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))))=1) then cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money) else 0.00 end +
'case When(isnumeric(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))))=1) then cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money) else 0.00 end 

cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money)+
cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) +
cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money)+
cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money)+
cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money)+
cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money)+
cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money)+
cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money)+
cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money)+
cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money)+
cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money)+
cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money)+
cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money)+
cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money)+
cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money)+
cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money)+
cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money)+
cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money)+
cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money)+
cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money)+
cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money)+
cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money)+
cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money)+
cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money)
,101) + ' not matching Details ' +
CONVERT(varchar,
isnull(
(select 
SUM(
cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money)+
cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) +
cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money)+
cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money)+
cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money)+
cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money)+
cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money)+
cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money)+
cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money)+
cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money)+
cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money)+
cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money)+
cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money)+
cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money)+
cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money)+
cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money)+
cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money)+
cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money)+
cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money)+
cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money)+
cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money)+
cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money)+
cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money)+
cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money))
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and dbo.fngetday(l2.dateprocessed)=dbo.fngetday(letterrequest.dateprocessed)
and l2.LetterCode='31005'),0.00),101)
from letterrequest with(nolock)
join master with(nolock) on master.number=letterrequest.accountid
where LetterCode in (31000, 31002, 31006)
--and DateRequested>=DATEADD(day,-1,{fn curdate()})
and Dateprocessed between @start and dateadd(ss,-1,dateadd(d,1,@start)) -->={fn curdate()}--DATEADD(day,-1,{fn curdate()})
--and YEAR(dateprocessed)=1753
and Deleted=0
--the total on the cover does not match the sum of the details
and 
NOT(
cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money)+
cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) +
cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money)+
cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money)+
cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money)+
cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money)+
cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money)+
cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money)+
cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money)+
cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money)+
cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money)+
cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money)+
cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money)+
cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money)+
cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money)+
cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money)+
cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money)+
cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money)+
cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money)+
cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money)+
cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money)+
cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money)+
cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money)+
cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money)
-

 isnull((select 
SUM(
cast(ltrim(RIGHT(sifpmt1,len(sifpmt1)-charindex(' ',sifpmt1))) as Money)+
cast(ltrim(RIGHT(sifpmt2,len(sifpmt2)-charindex(' ',sifpmt2))) as Money) +
cast(ltrim(RIGHT(sifpmt3,len(sifpmt3)-charindex(' ',sifpmt3))) as Money)+
cast(ltrim(RIGHT(sifpmt4,len(sifpmt4)-charindex(' ',sifpmt4))) as Money)+
cast(ltrim(RIGHT(sifpmt5,len(sifpmt5)-charindex(' ',sifpmt5))) as Money)+
cast(ltrim(RIGHT(sifpmt6,len(sifpmt6)-charindex(' ',sifpmt6))) as Money)+
cast(ltrim(RIGHT(sifpmt7,len(sifpmt7)-charindex(' ',sifpmt7))) as Money)+
cast(ltrim(RIGHT(sifpmt8,len(sifpmt8)-charindex(' ',sifpmt8))) as Money)+
cast(ltrim(RIGHT(sifpmt9,len(sifpmt9)-charindex(' ',sifpmt9))) as Money)+
cast(ltrim(RIGHT(sifpmt10,len(sifpmt10)-charindex(' ',sifpmt10))) as Money)+
cast(ltrim(RIGHT(sifpmt11,len(sifpmt11)-charindex(' ',sifpmt11))) as Money)+
cast(ltrim(RIGHT(sifpmt12,len(sifpmt12)-charindex(' ',sifpmt12))) as Money)+
cast(ltrim(RIGHT(sifpmt13,len(sifpmt13)-charindex(' ',sifpmt13))) as Money)+
cast(ltrim(RIGHT(sifpmt14,len(sifpmt14)-charindex(' ',sifpmt14))) as Money)+
cast(ltrim(RIGHT(sifpmt15,len(sifpmt15)-charindex(' ',sifpmt15))) as Money)+
cast(ltrim(RIGHT(sifpmt16,len(sifpmt16)-charindex(' ',sifpmt16))) as Money)+
cast(ltrim(RIGHT(sifpmt17,len(sifpmt17)-charindex(' ',sifpmt17))) as Money)+
cast(ltrim(RIGHT(sifpmt18,len(sifpmt18)-charindex(' ',sifpmt18))) as Money)+
cast(ltrim(RIGHT(sifpmt19,len(sifpmt19)-charindex(' ',sifpmt19))) as Money)+
cast(ltrim(RIGHT(sifpmt20,len(sifpmt20)-charindex(' ',sifpmt20))) as Money)+
cast(ltrim(RIGHT(sifpmt21,len(sifpmt21)-charindex(' ',sifpmt21))) as Money)+
cast(ltrim(RIGHT(sifpmt22,len(sifpmt22)-charindex(' ',sifpmt22))) as Money)+
cast(ltrim(RIGHT(sifpmt23,len(sifpmt23)-charindex(' ',sifpmt23))) as Money)+
cast(ltrim(RIGHT(sifpmt24,len(sifpmt24)-charindex(' ',sifpmt24))) as Money))
from letterrequest l2 with(nolock)
join master m2 with(nolock) on m2.number=l2.accountid
where m2.link=master.link and dbo.fngetday(l2.DateRequested)=dbo.fngetday(letterrequest.DateRequested)
and l2.deleted=0
and dbo.fngetday(l2.dateprocessed)=dbo.fngetday(letterrequest.dateprocessed)
and l2.LetterCode='31005'),0.00)
between -1 and 1)
and not exists(select * from @report r where r.number=accountid)

--link is now 0
--any of the covers or details letters have a bad address
--any of the covers or details have letters suppressed

insert into @report(number,lettercode,daterequested,dateprocessed,comment)
select letterrequest.AccountID,letterrequest.LetterCode,dbo.fngetday(letterrequest.DateRequested),dateprocessed,
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
--and DateRequested>=DATEADD(day,-1,{fn curdate()})
and Dateprocessed between @start and dateadd(ss,-1,dateadd(d,1,@start)) -->={fn curdate()}--DATEADD(day,-1,{fn curdate()})
--and YEAR(dateprocessed)=1753
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
order by master.link,r.number


end"
        End With

    End Sub

End Class