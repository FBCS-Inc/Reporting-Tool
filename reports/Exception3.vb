
Public Class Exception3

    Private Sub Exception3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Operations Exceptions 3 - Inventory Work Gap"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
            .shfmt = "TTCTTTTDTDD"
            .vsql = "declare @workDays int
            set @workDays=10
            Declare @contactDays int
            Set @contactDays=8

            Declare @accounts table(
            number int Not null,
            link int null,
            firstworked datetime null
            )

            insert into @accounts(number,link,firstworked)
            Select master.number,master.link,(
            Case when link=0 Or link Is null then 
            --this accounts first work date
            isnull(
            (select top 1 created from notes with(nolock) 
            join action With(nolock) On action.code=notes.action
            join result With(nolock) On result.code=notes.result
            where notes.number=master.number
            And (action.WasWorked=1 Or result.worked=1) 
            order by created)
            ,'1900-01-01') 
            Else
            --this accounts first work date after this was received
            isnull((
            Select top 1 n2.created
            from master m2 With(nolock)
            join notes  n2 With(nolock) On n2.number=m2.number
            join action a2 With(nolock) On a2.code=n2.action
            join result r2 With(nolock) On r2.code=n2.result
            where m2.link=master.link 
            And n2.created>=master.received
            And (a2.WasWorked=1 Or r2.worked=1) 
            order by n2.created desc),'1900-01-01') 
            End) 
            from master With(nolock)
            join debtors With(nolock) On debtors.Number=master.number And debtors.Seq=0
            join customer With(nolock) On customer.customer=master.customer
            join desk With(nolock) On desk.code=master.desk
            INNER JOIN [miscextra] score With (NOLOCK)
            On (score.[number] = [master].[number] And (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
            join Fact With(nolock) On fact.CustomerID=master.customer
            join CustomCustGroups opsgroup With(nolock) On opsgroup.ID=fact.CustomGroupID And opsgroup.Name Like 'ops%'
            where  
            --Placed between 4 And 30 days ago
            received between @start And
            DATEADD(day,-1*(case when customer.customvalue3='8' then 5 else 3 end+1),@start) 
            --its been at least 3 Or 5 business days since placement
            And (
            DATEDIFF(d,master.received,@start) 
            - DATEDIFF(wk,master.received,@start) * 2 
            - CASE WHEN DATENAME(dw, master.received) <> 'Saturday' AND DATENAME(dw, @start) = 'Saturday' THEN 1 
            WHEN DATENAME(dw, master.received) = 'Saturday' AND DATENAME(dw, @start) <> 'Saturday' THEN -1 
            Else 0 
            End ) 
            >= case when customer.customvalue3='8' then 5 else 3 end
            --Branch Is Not 2
            And desk.Branch<>'00002'
            --Open Status
            And closed Is null
            --No restrictions
            And (Not exists(select * from bankruptcy with(nolock) where bankruptcy.accountid=master.number And bankruptcy.datefiled <> '1753-01-01 12:00:00.000' and bankruptcy.datefiled<>'1900-01-01 00:00:00.000') and not exists(select * from deceased with(nolock) where deceased.accountid=master.number and deceased.dod is not null and deceased.dod <> '1753-01-01 12:00:00.000' and deceased.dod <> '1900-01-01 00:00:00.000') and not exists (select * from debtorattorneys with(nolock) where debtorattorneys.accountid=master.number and ( rtrim(debtorattorneys.[name]) <> '' or rtrim(debtorattorneys.[firm]) <> '' or rtrim(debtorattorneys.Addr1) <> '' or rtrim(debtorattorneys.Addr2) <> '' or rtrim(debtorattorneys.City) <> '' or rtrim(debtorattorneys.[State]) <> '' or rtrim(debtorattorneys.Zipcode) <> '' or rtrim(debtorattorneys.Phone) <> '' or rtrim(debtorattorneys.Fax) <> '' or rtrim(debtorattorneys.Email) <> '' or rtrim(debtorattorneys.Comments) <> '') ) and not exists(select * from restrictions with(nolock) where restrictions.number=master.number and ((not(home is null) and home <> 0) OR (not(job is null) and job <> 0)  OR  (not(calls is null) and calls <> 0) OR (not(comment is null) and not(comment like '')) OR (not(AttyName is null) and AttyName <> '') OR (not(Attyaddr1 is null) and Attyaddr1 <> '') OR (not(Attyaddr2 is null) and Attyaddr2 <> '') OR (not(AttyCity is null) and AttyCity <> '') OR (not(AttyState is null) and AttyState <> '') OR (not(AttyZip is null) and AttyZip <> '') OR (not(AttyPhone is null) and AttyPhone <> '') OR (not(Attynotes is null) and not(Attynotes like ''))OR (not(BkyCase is null) and BkyCase <> '') OR (not(BkyChap11 is null) and BkyChap11 <> 0) OR (not(BkyChap13 is null) and BkyChap13 <> 0) OR (not(BkyChap7 is null) and BkyChap7 <> 0) OR (not(BkyCourt is null) and BkyCourt <> '') OR (not(BkyDateFiled is null)) OR (not(BkyDistrict is null) and BkyDistrict <> '') OR (not(BkyNotes is null) and not(BkyNotes like '')) OR (not(suppressletters is null) and suppressletters <> 0) OR (not(Disputed is null) and Disputed <> 0))))
            --Has a phone in good Or unknown status
            And exists(select* from Phones_Master with(nolock) 
            where phones_master.Number=master.number 
            And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0
            And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2))
            --no attempt in 3 (regular) Or 5 days (medical) after placement
            And Not exists(select * from notes with(nolock) 
            join action With(nolock) On action.code=notes.action
            join result With(nolock) On result.code=notes.result
            where notes.number=master.number
            And (action.WasWorked=1 Or result.worked=1)
            And --datediff(day,master.received,notes.created)
            (
            DATEDIFF(d,master.received,notes.created) --Get the number of days between start And end dates
            - DATEDIFF(wk,master.received,notes.created) * 2 -- for each week, subtract 2 days (by default a week occurs between sat And sunday on sql server)
            - CASE WHEN DATENAME(dw, master.received) <> 'Saturday' AND DATENAME(dw, notes.created) = 'Saturday' THEN 1 --subtract 1 day if the end date falls on a saturday and the startdate is a weekday, or sunday --written by pflangan haha
            WHEN DATENAME(dw, master.received) = 'Saturday' AND DATENAME(dw, notes.created) <> 'Saturday' THEN -1 --add 1 if the start date is a saturday and the end date is any other day
            Else 0 
            End ) 
            <=
            Case when customer.customvalue3='8' then 5 else 3 end
            )

            --No Active Arrangements
            And Not exists(select * from PDC_View_FBCS pdc with(nolock) where pdc.number=master.number And pdc.Active=1)
            And Not exists(select * from PDCC_View_FBCS pdc with(nolock) where pdc.number=master.number And pdc.isActive=1)
            And Not exists(select * from promises with(nolock) where promises.acctid=master.number And promises.Active=1)
            And Not exists(select * from paymentbatchitems p with(nolock) where p.FileNum=master.number And p.PmtType=1)
            --Last Good Posititve Payment was more than 30 days ago
            And isnull((select top 1 entered from payhistory p1 with(nolock) 
            	where p1.number=master.number And p1.batchtype In ('PU','PC') 
            	And Not exists(select * from payhistory p1rev with(nolock) where p1rev.number=master.number
            	And p1rev.batchtype in ('PUR','PCR')
            	And p1rev.ReverseOfUID=p1.UID)
            	order by entered desc),'1-1-1900')
            <DATEADD(Day, -30,@start)
            --Exclude desks Spanish, Spanish14, A0115, PENDCLSE & All Collector Desks
            And master.desk Not in ('Spanish', 'Spanish14', 'A0115', 'PENDCLOSE')
            And desk.desktype Not in ('Collector')
            --Also exclude clients MCM & Turnstile
            And opsgroup.name Not Like '%ops%midland%'
            And opsgroup.name Not Like '%ops%turnstile%'

            Select top 2500 master.number [File #], opsgroup.name [Client Ops Group],
            master.current0 [Balance],master.State [State],
            score.thedata [Score],master.Desk + ': ' + desk.name [Desk],
            Case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2) And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0 And phones_master.PhoneTypeID=1) then 'Home ' else '' end
            +
            Case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2) And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0 And phones_master.PhoneTypeID=2) then 'Work ' else '' end
            +
            Case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2) And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0 And phones_master.PhoneTypeID=3) then 'Cell ' else '' end
            +
            Case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2) And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0 And phones_master.PhoneTypeID=7) then 'Skip ' else '' end
            +
            Case when exists(select * from Phones_Master with(nolock) where phones_master.Number=master.number And (phones_master.PhoneStatusID Is null Or phones_master.PhoneStatusID=2) And phones_master.debtorid=debtors.debtorid And phones_master.OnHold=0 And phones_master.PhoneTypeID Not in (1,2,3,7)) then 'Other ' else '' end
            [Phone Types],
            convert(varchar,master.Received,101) [Received],master.Status,
            isnull(convert(varchar,master.worked,101),'') [Last Worked], case when a.firstworked='1-1-1900' then '' else convert(varchar,a.firstworked,101) end [First Worked] 
            from @accounts a
            join master With(nolock) On master.number=a.number
            join debtors With(nolock) On debtors.Number=master.number And debtors.Seq=0
            join customer With(nolock) On customer.customer=master.customer
            join desk With(nolock) On desk.code=master.desk
            INNER JOIN [miscextra] score With (NOLOCK)
            On (score.[number] = [master].[number] And (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
            join Fact With(nolock) On fact.CustomerID=master.customer
            join CustomCustGroups opsgroup With(nolock) On opsgroup.ID=fact.CustomGroupID And opsgroup.Name Like 'ops%'
            where  
            (
            a.firstworked='1-1-1900'
            Or
            a.firstworked Is null
            Or
            (
            DATEDIFF(d,master.received,a.firstworked) --Get the number of days between start And end dates
            - DATEDIFF(wk,master.received,a.firstworked) * 2 -- for each week, subtract 2 days (by default a week occurs between sat And sunday on sql server)
            - CASE WHEN DATENAME(dw, master.received) <> 'Saturday' AND DATENAME(dw, a.firstworked) = 'Saturday' THEN 1 --subtract 1 day if the end date falls on a saturday and the startdate is a weekday, or sunday --written by pflangan haha
            WHEN DATENAME(dw, master.received) = 'Saturday' AND DATENAME(dw, a.firstworked) <> 'Saturday' THEN -1 --add 1 if the start date is a saturday and the end date is any other day
            Else 0 
            End ) 
            >
            Case when customer.customvalue3='8' then 5 else 3 end
            ) order by master.received desc"
        End With
    End Sub

End Class