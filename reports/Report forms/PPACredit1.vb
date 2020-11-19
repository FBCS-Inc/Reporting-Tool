Public Class PPACredit1
	Private Sub PPACredit1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "PPA Credit Exceptions"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "TTT"
			.vsql = "declare @uid integer
set @uid=(select MAX(uid) from PPACreditExceptions with(nolock))
declare @date1 datetime;
declare @date2 datetime;
declare @date3 datetime;
set @date1 = DATEADD(day,-1*(DAY(@start))+1,@start)
set @date2 = DATEADD(day,-15,@start)
DROP TABLE IF EXISTS tempdb.dbo.#tempNotes
DROP TABLE IF EXISTS tempdb.dbo.#number1
DROP TABLE IF EXISTS tempdb.dbo.#exclude1
DROP TABLE IF EXISTS tempdb.dbo.#include1
DROP TABLE IF EXISTS tempdb.dbo.#include2
DROP TABLE IF EXISTS tempdb.dbo.#include3
Create table dbo.#tempNotes(number int not null,error1 varchar(200) not null,error2 varchar(200) not null)
CREATE CLUSTERED INDEX idx_tempNotes_Number ON #tempNotes(Number)
Create table dbo.#number1(number int not null,entered datetime null)
CREATE CLUSTERED INDEX idx_number1_Number ON #number1(Number)
Create table dbo.#exclude1(number int not null)
CREATE CLUSTERED INDEX idx_exclue1_Number ON #exclude1(Number)
Create table dbo.#include1(number int not null)
CREATE CLUSTERED INDEX idx_include1_Number ON #include1(Number)
Create table dbo.#include2(number int not null)
CREATE CLUSTERED INDEX idx_include2_Number ON #include2(Number)
Create table dbo.#include3(number int not null)
CREATE CLUSTERED INDEX idx_include3_Number ON #include3(Number)
exec nowait_print 'Starting'
insert into #tempNotes(number,error1,error2)
select payhistory.Number,
creddesk.TheData + ' is not a valid PPACreditDesk',
(select top 1 p2.desk + ' - ' + d2.name from payhistory p2 with(nolock)
join desk d2 with(nolock) on d2.code=p2.desk
where p2.number=payhistory.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%' order by p2.entered desc) [Last Collector Payment Desk]
from payhistory with(nolock)
join desk with(nolock) on desk.code=payhistory.desk
join miscextra creddesk with(nolock) on creddesk.Number=payhistory.Number and creddesk.Title like '%PPa%credit%desk%'
where desk.Branch='00002'
and batchtype in ('PU','PC')
and not exists(select * from desk d1 with(nolock) where d1.code=creddesk.TheData)
and payhistory.entered >= @date1
and RTRIM(creddesk.TheData)<>''
and RTRIM(creddesk.TheData)<>'0'
and exists(select * from payhistory p2 with(nolock)
where p2.number=payhistory.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%')
exec nowait_print 'Insert 1 finished'
insert into #tempNotes(number,error1,error2)
select payhistory.Number,
creddate.TheData + ' is not a valid PPACreditDate',
(select top 1 p2.desk + ' - ' + d2.name from payhistory p2 with(nolock)
join desk d2 with(nolock) on d2.code=p2.desk
where p2.number=payhistory.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%' order by p2.entered desc) [Last Collector Payment Desk]
from payhistory with(nolock)
join desk with(nolock) on desk.code=payhistory.desk
join miscextra creddate with(nolock) on creddate.Number=payhistory.Number and creddate.Title = 'PPACreditDate'
where desk.Branch='00002'
and batchtype in ('PU','PC')
and RTRIM(creddate.TheData)<>''
and RTRIM(creddate.TheData)<>'0'
and ISDATE(creddate.thedata)=0
and payhistory.entered >= @date1
and exists(select * from payhistory p2 with(nolock)
where p2.number=payhistory.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%')
exec nowait_print 'Insert 2 finished'
insert into #tempNotes(number,error1,error2)
select miscextra.Number,
case 
when miscextra.title <> 'PPACreditDate' then 'PPACreditDate is not spelled correctly'
when isdate(miscextra.thedata)=0 then 'PPACreditDate is not a valid date '  + RTRIM(miscextra.thedata)
else '' end,''
from notes with(nolock)
join miscextra with(nolock) on miscextra.Number=notes.number 
and miscextra.Title like '%PPA%CREDIT%DATE%'
and RTRIM(miscextra.thedata)<>''
where created>=@date2
and notes.comment like '%misc%extra%ppa%credit%date%'
and not exists(select * from miscextra mcheck with(nolock) where mcheck.Number=miscextra.Number
and mcheck.Title='PPACreditDate'
and ISDATE(mcheck.thedata)=1)
exec nowait_print 'Insert 3 finished'
insert into #tempNotes(number,error1,error2)
select miscextra.Number,'PPACreditDesk is not spelled correctly',''
from notes with(nolock)
join miscextra with(nolock) on miscextra.Number=notes.number 
and miscextra.Title like '%PPA%CREDIT%DESK%'
and RTRIM(miscextra.thedata)<>''
where created>=@date2
and Title not like 'PPA%Credit%Desk'
and notes.comment like '%misc%extra%ppa%credit%desk%'
and not exists(select * from miscextra mcheck with(nolock) where mcheck.Number=miscextra.Number
and mcheck.Title like 'PPA%Credit%Desk')
exec nowait_print 'Insert 4 finished'
--Bad value in ppa credit desk
insert into #tempNotes(number,error1,error2)
select miscextra.Number,
miscextra.TheData + ' is not a valid PPACreditDesk',
(select top 1 p2.desk + ' - ' + d2.name from payhistory p2 with(nolock)
join desk d2 with(nolock) on d2.code=p2.desk
where p2.number=notes.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%' order by p2.entered desc) [Last Collector Payment Desk]
from notes with(nolock)
join miscextra with(nolock) on miscextra.Number=notes.number 
and miscextra.Title like '%PPA%CREDIT%DESK%'
where created>=@date2
and RTRIM(miscextra.TheData)<>''
and RTRIM(miscextra.TheData)<>'0'
and notes.comment like '%misc%extra%ppa%credit%desk%'
and not exists(select * from desk d1 with(nolock) where d1.code=miscextra.TheData)
and exists(select * from payhistory p2 with(nolock)
where p2.number=notes.number
and p2.batchtype in ('PU','PC')
and p2.desk like 'A0%')
exec nowait_print 'Insert 5 finished'
insert into #number1 (number,entered)
SELECT payhistory.number, payhistory.entered
FROM   payhistory WITH(nolock) 
       JOIN desk WITH(nolock)  ON desk.code = payhistory.desk 
WHERE  desk.branch = '00002' 
       AND batchtype IN ( 'PU', 'PC' ) 
       AND payhistory.entered >= @date1 
insert into #exclude1 (number)
       SELECT n1.number 
        FROM   miscextra creddesk WITH(nolock) join #number1 n1 on creddesk.number = n1.number
        WHERE  creddesk.number = n1.number 
                AND creddesk.title LIKE '%PPa%credit%desk%' 
insert into #exclude1 (number)
       SELECT n1.number               
        FROM   miscextra creddesk WITH(nolock) join #number1 n1 on creddesk.number = n1.number
        WHERE  creddesk.number = n1.number 
                AND creddesk.title LIKE 'PPACreditDate'
delete from #number1 where number in (select number from #exclude1);
insert into #include1 (number)
SELECT n1.number
FROM   payhistory p2 WITH(nolock) join #number1 n1 on p2.number = n1.number 
WHERE  p2.number = n1.number 
        AND p2.batchtype IN ( 'PU', 'PC' ) 
        AND p2.desk LIKE 'A0%' 
insert into #include2 (number)
SELECT n1.number
FROM   notes n WITH(nolock) join #number1 n1 on n.number = n1.number
WHERE  n.number = n1.number 
        AND n.comment LIKE 'misc extra%ppa%term%' 
        AND n.created < n1.entered
insert into #include3 (number)
select number from #number1 n1
where (SELECT Count(*) 
            FROM   payhistory p2 WITH(nolock) 
            WHERE  p2.number = n1.number 
                   AND p2.batchtype IN ( 'PU', 'PC' )) = 1 
delete from #number1 
where number not in 
(select i1.number from #include1 i1 join #include2 i2 on i1.number = i2.number
								join #include3 on i1.number = i2.number)
--had a pay to a coll desk and ppacredit desk is missing all together 
INSERT INTO #tempnotes 
            (number, 
             error1, 
             error2) 
SELECT payhistory.number, 
       CASE 
         WHEN NOT EXISTS(SELECT * 
                         FROM   miscextra creddesk WITH(nolock) 
                         WHERE  creddesk.number = payhistory.number 
                                AND creddesk.title LIKE 'PPACreditDate') THEN 
         'Missing PPACreditDate' 
         ELSE 'Missing PPACreditDesk' 
       END, 
       (SELECT TOP 1 p2.desk + ' - ' + d2.NAME 
        FROM   payhistory p2 WITH(nolock) 
               JOIN desk d2 WITH(nolock) 
                 ON d2.code = p2.desk 
        WHERE  p2.number = payhistory.number 
               AND p2.batchtype IN ( 'PU', 'PC' ) 
               AND p2.desk LIKE 'A0%' 
        ORDER  BY p2.entered DESC) [Last Collector Payment Desk] 
FROM   payhistory WITH(nolock) 
       JOIN desk WITH(nolock) 
         ON desk.code = payhistory.desk 
		 join #number1 n1 on payhistory.number = n1.number and payhistory.entered = n1.entered
exec nowait_print 'Inserts finished'
insert into PPACreditExceptions(number,exception,lastcollectorpaydesk)
select distinct 
number [File Number],
error1 as [Reason],
error2 as [Last Collector Payment Desk]
from #tempNotes
select distinct number [File Number], 
substring(exception,0,40) as [Reason], 
substring(lastcollectorpaydesk,0,75) as [Last Collector Payment Desk]
from PPACreditExceptions with(nolock) 
where [uid] >  cast(@uid as varchar)
 order by number"
		End With
	End Sub
End Class