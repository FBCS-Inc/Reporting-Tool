

Public Class Settlement
    Private Sub Settlement_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Settlements to be 2nd reviewed"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
            .shfmt = "TTTTDCCTCCC"
            .vsql = "Select distinct customcustgroups.Name [Ops Group Name],customer.customer + ': ' + customer.name [Customer], "
            .vsql += " master.number [File Number],master.Desk,Convert(varchar, payhistory.entered, 101) [Last Paid], "
            .vsql += " master.Original [Original Balance],master.current0 [Current Balance],Case when customer.BlanketSif between 1.0 And 99.9 then "
            .vsql += " cast(customer.blanketsif As varchar) Else isnull(BlanketSif.thedata,'100') end [BlanketSif], "
            .vsql += " Convert(varchar,case when ISNUMERIC(sifamount.thedata) = 1 Then sifamount.thedata Else '0.00' end,101) [SifAmount],"
            .vsql += " -paid [Total Paid], cast(case  when customer.BlanketSif between 1.0 And 99.9 then  cast(customer.blanketsif As varchar) "
            .vsql += " when ISNUMERIC(BlanketSif.thedata)=0 then 100 Else isnull(BlanketSif.thedata,100) end as Money) / 100.0 * master.original  [Blanket Sif Amount]"
            .vsql += " From master with(nolock)"
            .vsql += " Join status with(nolock) on status.code=master.status "
            .vsql += " Join customer with(nolock) on customer.customer=master.customer"
            .vsql += " Join Fact with(nolock) on fact.CustomerID=master.customer"
            .vsql += " Join customcustgroups with(nolock) on customcustgroups.ID=fact.CustomGroupID And customcustgroups.Name Like 'ops%'"
            .vsql += " Left outer join miscextra BlanketSif with(nolock) on BlanketSif.Number=master.number And BlanketSif.Title='BlanketSif' "
            .vsql += " Join miscextra sifamount with(nolock) on sifamount.Number=master.number And sifamount.Title='SifAmount' And ISNUMERIC(sifamount.thedata)=1 "
            .vsql += " Join payhistory with(nolock) on payhistory.UID=(Select top 1 p.uid from payhistory p with(nolock) 	where p.number = master.number And p.batchtype in ('PU','PC','PUR','PCR') 	order by p.entered desc, p.UID desc) "
            .vsql += " where -paid > 0 "
            .vsql += " And CAST(case when isnumeric(sifamount.thedata)=1 then sifamount.thedata else '0.00' end as money) > 0"
            .vsql += " And payhistory.batchtype in ('PU','PC')"
            .vsql += " And payhistory.entered between @start And @end"
            .vsql += " And -paid >= CAST(case when isnumeric(sifamount.thedata)=1 then sifamount.thedata else '0.00' end as money)-.05"
            .vsql += " And closed Is null"
            .vsql += " And status.statustype Like '0%'"
            .vsql += " And master.desk Not in ('SIF14','SIF21', 'SIF30', 'PIF15', 'PIF21', 'PIF30', 'SIFPENDNSF', 'UBS','HOLD14POST')"
            .vsql += " And Not exists(select * from PDC_View_FBCS pdc with(nolock) where pdc.number=master.number And pdc.Active=1)"
            .vsql += " And Not exists(select * from PDCC_View_FBCS pdcc with(nolock) where pdcc.number=master.number And pdcc.isActive=1)"
            .vsql += " And Not exists(select * from Promises_View_FBCS prom with(nolock) where prom.AcctID=master.number And prom.Active=1 And (prom.Suspended Is null Or prom.Suspended=0))"
            .vsql += " And ((case when customer.BlanketSif between 1.0 And 99.9 then cast(customer.blanketsif as varchar) else isnull(BlanketSif.thedata,'100') end = '100')"
            .vsql += " Or (case  when customer.BlanketSif between 1.0 And 99.9 then cast(customer.blanketsif as varchar) when ISNUMERIC(BlanketSif.thedata)=0 then '100' else isnull(BlanketSif.thedata,'100') end <> '100'"
            .vsql += " And ROUND(-paid,2) >= round(cast(case when customer.BlanketSif between 1.0 And 99.9 then cast(customer.blanketsif as varchar)when ISNUMERIC(BlanketSif.thedata)=0 then '100' else isnull(BlanketSif.thedata,'100') end as Money)/100.0 * master.original ,2)-5.00))"
            .Refresh()
            .Show()

        End With


    End Sub


End Class