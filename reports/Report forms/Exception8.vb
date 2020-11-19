Public Class Exception8
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        With ReportData1
            .RPTNAME = "Operational Exceptions 8 - Timezone"
            .FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte+.ext
            .shfmt = "TTCTTTTDTTDDDT"
            .vsql = "declare @report table(
number int not null,
phone varchar(30) not null,
ESTCalldatetime datetime not null,
ActualCallDateTime_FBCS datetime not null,
ActualCallDateTime_Livevox datetime not null,
additionalInfo varchar(200))

insert into @report(number,phone,ESTCalldatetime,ActualCallDateTime_FBCS,ActualCallDateTime_Livevox,additionalInfo)
select campaignresults.AccountID,
phonenum,
CampaignResults.CallDateTime [EST Call Time],
dbo.[FBCS_ConvertToTimeZone](livevoxresults.CallDateTime,FBCSAreaCodeList_2.timezone)
 [Consumer's Call Time - FBCS Calculated],
CampaignResults.ActualCallDateAndTime [Consumer's Call Time - Livevox Calculate],
case when FBCSAreaCodeList_2.areacode is null then 'Unmapped area code'
else '' end
from campaignresults with(nolock)
join livevoxresults with(nolock) on livevoxresults.accountid=CampaignResults.AccountID and livevoxresults.CallDateTime=campaignresults.CallDateTime
left outer join FBCSAreaCodeList_2 with(nolock) on FBCSAreaCodeList_2.areacode=LEFT(livevoxresults.phonenum,3)
left outer join states with(nolock) on states.Description=FBCSAreaCodeList_2.state
where livevoxresults.CallDateTime>=DATEADD(day,-7,@start)
--calldate time in the area code's timezone
and (
(FBCSAreaCodeList_2.areacode is null
and livevoxresults.result not in ('DC')
and livevoxresults.phonenum not like '8%'
)
or
datepart(HH,dbo.[FBCS_ConvertToTimeZone](livevoxresults.CallDateTime,FBCSAreaCodeList_2.timezone))
not between
	case
	when rtrim(ltrim(states.code)) in ('LA') then 8
	when rtrim(ltrim(states.code)) in ('CA','ID','IL','MN','MS','OK','PA','TX','NV')
	then 9
	else 8 end
and 
	case 
	when states.code like 'LA%' then 20
	when states.code like 'NV%' then 20
	else 21 end
)


select master.number [File #], opsgroup.name [Client Ops Group],
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
a.phone [Phone Dialed],a.ActualCallDateTime_FBCS,a.ActualCallDateTime_Livevox,a.ESTCalldatetime,
a.additionalInfo
from @report a
join master with(nolock) on master.number=a.number
join debtors with(nolock) on debtors.Number=master.number and debtors.Seq=0
join customer with(nolock) on customer.customer=master.customer
join desk with(nolock) on desk.code=master.desk
INNER JOIN [miscextra] score WITH (NOLOCK)
ON (score.[number] = [master].[number] and (score.[title] = 'TU_MiscInfo1' or score.[title] = 'TUMiscInfo1'))
join Fact with(nolock) on fact.CustomerID=master.customer
join CustomCustGroups opsgroup with(nolock) on opsgroup.ID=fact.CustomGroupID and opsgroup.Name like 'ops%'
order by master.received desc
"
        End With
    End Sub

End Class