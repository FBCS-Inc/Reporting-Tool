Public Class TSISMERestrictionRemoval
	Dim sql As String
	Private Sub TSISMERestrictionRemoval_Load(sender As Object, e As EventArgs) Handles MyBase.Load


	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		sql = "--Items 1 through 5 below may or may not need to be updated, but should be checked.  
--Item #6 will always need to updated

--Use this script to release the restrictions for the zip table that was requested.
--Once you are complete running this script comment out the zip table criteria in the Disaster Holds job for the zip 
--table we are no longer restricting
  


SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED

declare @metitle varchar(50)

DECLARE @number INTEGER
DECLARE @debtor INTEGER
DECLARE @seq INTEGER
DECLARE @rid INTEGER
DECLARE @comment VARCHAR(500)
declare @disaster varchar(150)
declare @disasterdate datetime
declare @desk varchar(15)
--1,2,3,
DECLARE @accounts CURSOR
SET @accounts = CURSOR FAST_FORWARD READ_ONLY FOR
select distinct  master.number,debtors.debtorid, 
restrictions.RestrictionID , cast(restrictions.comment as varchar(1000)),
  'Pandemic Hold' ----1. Here modify for the name used for the disaster, (Look at the Disaster Hold job and see what the disaster name variable was assigned when adding the restrictions)
 ------------------------currently it is looking for disaster ' Storms / Flooding' which is what we've been using
,
nresAdded.created,master.desk, debtors.seq
from master with(nolock)
join debtors with(nolock) on debtors.Number=master.number --and debtors.seq=1
join status with(nolock) on status.code=master.status
join miscextra with(nolock) 
on ( miscextra.number=master.number and title=
'ClientTempHoldDate'
+
case when debtors.Seq>0 then '_' + cast(debtors.Seq+1 as varchar) else '' end
)

join restrictions with(nolock) on (restrictions.debtorid=debtors.debtorid)
join customer with (nolock) on (master.customer = customer.customer )
join Fact with (nolock) on (fact.CustomerID = customer.customer )
join CustomCustGroups with (nolock) on (CustomCustGroups.ID = fact.CustomGroupID 
and CustomCustGroups.Name like 'FIN/%')

join notes nresAdded with(nolock)  on
	nresAdded.UID=(select top 1 n1.uid from notes n1 with(nolock)
	where n1.number=master.number 
	and n1.comment like 'Debtor%' + CAST(debtors.seq+1 as varchar) + '%'  
	and n1.comment like '%Restrictions Added due to%Pandemic Hold%' 
	and n1.comment not like 'Restrictions Added due to%release%' 
	and n1.created>='02-01-2020'order by n1.created desc) --2. Here you would change date to when holds started
where 

--hold date is during the hurricane/disater hold period
case when ISDATE(TheData)=1 then
cast(TheData as datetime) 
else
cast('1-1-1900' as datetime) 
end 
between '02-01-2020' and {fn curdate()} + ' 23:59:59.00' --3. Here for start date you would change date to when holds started

and restrictions.suppressletters=1

--restriction updated during the hurricane/disater hold period
and restrictions.DateUpdated between '02-01-2020' and {fn curdate()} + ' 23:59:59.00' --4. Here for start date you would change date to when holds started

and exists(Select *  from notes nadd with(nolock)
where nadd.number=nresAdded.number and nadd.created>=dateadd(ss,-5,nresAdded.created)
and (
nadd.comment like  '%Restrictions Added due to%Pandemic Hold%'  --5. Here modify for the name used for the disaster, (Look at the Disaster Hold job and see what the disaster name variable was assigned when adding the restrictions)
----------------------------------------------02--------------------currently it is looking for disaster 'Storms/Flooding' which is what we've been using
--or nadd.comment like 'Restrictions Added due to Storm%'  
)
)

--these restrictions have not been released already
and not exists(Select *  from notes nrelease with(nolock)
where nrelease.number=nresAdded.number and nrelease.created>nresAdded.created
and nrelease.comment like 'Debtor%' + CAST(debtors.seq+1 as varchar) + '%'  
and nrelease.comment like '%Restrictions Added due to%release%'  )

--Don't want to release restrictions if the account is in one of the below desks
and master.desk not in ('COMP1','FDCPAHOLD','CEASE','RESOUTSTAT','ATTREP','HLDBNKOPND','DECREV','BANKOREV')



--these are test accounts so exclude them
and master.number not in (47001,47005)
--this is a test customer so exclude it
and customer.Name not like '%AARON%APPLIANCES%'


and CustomCustGroups.Name = 'FIN/TSI-SME ONLY'

---FIND the pattern used in JOB: Misc Sweeps: Disaster Holds
--STEP: TSI Pan Hold
-- for the creditors you need to release and put below
and  (
--Comprehensive Ortho&Muscoloskeletal Care
 originalcreditor like '" + textbox1.text + "'

--Lake Cook Orthopedics
--or originalcreditor like '%Lake%Cook%Orthopedic%'

)

 
 and debtors.state not like 'DC%'


--199829679
and master.number not in (47001,47005)
and customer.Name not like '%AARON%APPLIANCES%'








OPEN @accounts

FETCH NEXT FROM @accounts INTO @number, @debtor,@rid, @comment, @disaster, @disasterdate, @desk, @seq

WHILE @@FETCH_STATUS = 0 BEGIN

	BEGIN TRANSACTION
             
     print @number
     
     --only for inventory desks
     
     --per alexis we can lift letter restriction on edited restrictions on these desks:
     --'DOC','DISPUTES', 'PAIDPRIOR', 'FRAUD'
     
     if 
     --desk is NOT approved by alexis
     @desk not in ('DOC','DISPUTES', 'PAIDPRIOR', 'FRAUD') 
     --and the restrictions were edited after the hold
     and exists(select * from restrictions r with(nolock)
     where r.number=@number and r.debtorid=@debtor 
     and dbo.fngetday(r.DateUpdated) > dbo.fngetday(@disasterdate))
     begin
		--there is no co-debtor to explain the editing of restrictions - skip account
		if not exists(select * from debtors c with(nolock) where c.Number=@number and c.DebtorID<>@debtor)
		begin
			print cast(@number as varchar) + ' | skipping edited restrictions - no codebtor'
			ROLLBACK TRANSACTION
			GOTO NEXTLOOP
		
		end
		--a user edited restrictions - skip account
		if exists(select * from notes with(nolock) 
			where notes.number=@number and notes.created>@disasterdate 
			and notes.comment like 'Debtor(%) Restrict%Set%')
		begin
			print cast(@number as varchar) + ' | skipping edited restrictions - edited by user'
			ROLLBACK TRANSACTION
			GOTO NEXTLOOP
		end
		--a sweep edited restrictions - skip account
		if exists(select * from notes with(nolock) 
			where notes.number=@number and notes.created>@disasterdate 
			and notes.comment like '%Restrict%Set%')
		begin
			print cast(@number as varchar) + ' | skipping edited restrictions - edited by sweep'
			ROLLBACK TRANSACTION
			GOTO NEXTLOOP
		end
		
     end
     
     if @disaster is null or @disaster='' begin
		print cast(@number as varchar) + ' | unknown disater'
		ROLLBACK TRANSACTION
		GOTO NEXTLOOP
     end

	if @comment = 'ClientTempHold' begin

		update restrictions
		set restrictions.suppressletters=0,
		restrictions.dateupdated=getdate(), comment = null
		from restrictions with(rowlock)
		where debtorid=@debtor
		and number=@number and restrictionid = @rid
		
	end else begin
		update restrictions
		set restrictions.suppressletters=0,
		restrictions.dateupdated=getdate()
		from restrictions with(rowlock)
		where debtorid=@debtor
		and number=@number and restrictionid = @rid
	end

	IF NOT @@ERROR = 0 BEGIN
		ROLLBACK TRANSACTION
		GOTO NEXTLOOP
	END
	
	--Put Release Date in miscextra
	
	set @metitle='ClientTempHoldReleaseDate'
	if @seq > 0 begin 
		set @metitle = @metitle + '_' + cast(@seq+1 as varchar)
	end
	
	if exists(select * from miscextra with(nolock) where number=@number and title=@metitle) begin
		update miscextra
		set miscextra.thedata=convert(varchar,getdate(),101)
		from MiscExtra with(rowlock)
		where number=@number
		and title=@metitle
	end else begin
		insert into miscextra (number,title,thedata)
		values (@number,@metitle, convert(varchar,getdate(),101))
	end

--Document account released in the notes		
insert into notes 
(number, comment, action, result, created, user0)
values
(@number, 'Debtor(' + cast(@seq+1 as varchar) + ') Restrictions Added Due to ' + @disaster + ' Have Been Released', '+++++','+++++',getdate(), 'SYS')

	IF NOT @@ERROR = 0 BEGIN
		ROLLBACK TRANSACTION
		GOTO NEXTLOOP
	END
	
	COMMIT TRANSACTION

	NEXTLOOP:
	FETCH NEXT FROM @accounts INTO @number, @debtor,@rid, @comment, @disaster, @disasterdate, @desk, @seq
END


CLOSE @accounts
DEALLOCATE @accounts"
	End Sub
End Class