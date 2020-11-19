Public Class DebtSettlement
    Private Sub DebtSettlement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Debt Settlement Query"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
            .shfmt = "TTTTCTTTTTT"
            .StartDate.Visible = False
            .StartLabel.Visible = False
            .enddate.Visible = False
            .EndLabel.Visible = False
            .vsql = "select  master.customer [Client ID],
debtors.ssn [SSN],  
--substring(d.name,1,Charindex(',',d.Name)-1) [Last Name],
--substring(d.Name,Charindex(',',d.Name)+1,LEN(d.name)-(Charindex(',',d.Name))) [First Name],
debtors.name,
master.account [Original Account],
master.current0 [Balance], 
debtors.Street1 [Address 1],
debtors.Street2 [Address 2],
debtors.City [City],
debtors.State [State],
debtors.Zipcode [Zip],
CASE WHEN debtors.HomePhone <> '' then debtors.HomePhone
WHEN debtors.pager <> '' then debtors.pager
WHEN debtors.WorkPhone <> '' then debtors.WorkPhone  end
[Phone]

from master  with(nolock)
inner join Debtors  with(nolock)on debtors.number=master.number and debtors.Seq=0
where master.closed is null
AND [master].[desk] IN ('BRANCH1INV', 'BRANCH1VIR', 'BRANCH5INV', 'BRANCH5VIR', 'BRANCH5', 'BRANCH6INV', 
'BRANCH6VIR', 'DIALERSALV', 'EXPNOHIT','A0115', 'SALVAGE', 'SALVAGE2', 'SALVIN', 'SALVBRANCH','CLOSES')

AND [master].[status] IN ('ACT', 'HOT', 'NEW', 'NPC', 'NSF')
--in a collector queue
AND (([master].[qlevel] BETWEEN '000' AND '399')
     OR ([master].[qlevel] BETWEEN '400' AND '424')
     OR ([master].[qlevel] BETWEEN '425' AND '599'))
--no arrangements
and not exists(SELECT * FROM dbo.PDC_View_FBCS pdc with(nolock) where pdc.number=master.number and pdc.Active=1)
and not exists(SELECT * FROM dbo.PDCC_View_FBCS pdcc with(nolock) where pdcc.number=master.number and pdcc.IsActive=1)
and not exists(SELECT * FROM dbo.Promises_View_FBCS prom with(nolock) where prom.AcctID=master.number 
and prom.Active=1 and (prom.Suspended=0 or prom.Suspended is null))
and (DATEDIFF(day,lastpaid, {fn curdate()}) > 90 or lastpaid is null)

--no restrictions
AND (not exists(select * from bankruptcy with(nolock) where bankruptcy.accountid=master.number 
and bankruptcy.debtorid=debtors.debtorid and bankruptcy.datefiled is not null and bankruptcy.datefiled <> '1753-01-01 12:00:00.000' 
and bankruptcy.datefiled<>'1900-01-01 00:00:00.000') and not exists(select * from deceased with(nolock) where 
deceased.accountid=master.number and deceased.debtorid=debtors.debtorid and deceased.dod is not null 
and deceased.dod <> '1753-01-01 12:00:00.000' and deceased.dod <> '1900-01-01 00:00:00.000') 
and not exists(select * from debtorattorneys with(nolock) where debtorattorneys.accountid=master.number 
and debtorattorneys.debtorid=debtors.debtorid and ( rtrim(debtorattorneys.[name]) <> '' or rtrim(debtorattorneys.[firm]) <> '' or rtrim(debtorattorneys.Addr1) <> '' or rtrim(debtorattorneys.Addr2) <> '' or rtrim(debtorattorneys.City) <> '' or rtrim(debtorattorneys.[State]) <> '' or rtrim(debtorattorneys.Zipcode) <> '' or rtrim(debtorattorneys.Phone) <> '' or rtrim(debtorattorneys.Fax) <> '' or rtrim(debtorattorneys.Email) <> '' or rtrim(debtorattorneys.Comments) <> '') ) and not exists(select * from restrictions with(nolock) where restrictions.number=master.number and restrictions.debtorid=debtors.debtorid and ((not(home is null) and home <> 0) OR (not(job is null) and job <> 0)  OR  (not(calls is null) and calls <> 0) OR (not(comment is null) and not(comment like '')) OR (not(AttyName is null) and AttyName <> '') OR (not(Attyaddr1 is null) and Attyaddr1 <> '') OR (not(Attyaddr2 is null) and Attyaddr2 <> '') OR (not(AttyCity is null) and AttyCity <> '') OR (not(AttyState is null) and AttyState <> '') OR (not(AttyZip is null) and AttyZip <> '') OR (not(AttyPhone is null) and AttyPhone <> '') OR (not(Attynotes is null) and not(Attynotes like ''))OR (not(BkyCase is null) and BkyCase <> '') OR (not(BkyChap11 is null) and BkyChap11 <> 0) OR (not(BkyChap13 is null) and BkyChap13 <> 0) OR (not(BkyChap7 is null) and BkyChap7 <> 0) OR (not(BkyCourt is null) and BkyCourt <> '') OR (not(BkyDateFiled is null)) OR (not(BkyDistrict is null) and BkyDistrict <> '') OR (not(BkyNotes is null) and not(BkyNotes like '')) 
OR (not(suppressletters is null) and suppressletters <> 0) OR (not(Disputed is null) and Disputed <> 0))))"
        End With
    End Sub
End Class