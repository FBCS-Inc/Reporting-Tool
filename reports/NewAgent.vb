Public Class NewAgent
	Private Sub NewAgent_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		With ReportData1
			.enddate.Visible = False
			.EndLabel.Visible = False
			.RPTNAME = "New LiveVox Agent Found"
			.FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + .RPTNAME + .rptdte + .ext
			.shfmt = "T"
			.vsql = "select distinct agent
from livevoxinbounds li with(nolock)
where li.calldate>=@start
and li.agent not in (Select lu.AgentName from livevoxusers lu with(nolock) )
and li.agent<>''
union select distinct l.HCIAgent [agent]
from livevoxresults l with(nolock)
where l.CallDateTime>=DATEADD(day,-1,{fn curdate()})
and HCIAgent is not null 
and  HCIAgent<>''
and not exists(select * from livevoxusers u with(nolock)
where u.HCIAgentName=l.HCIAgent)"
		End With
	End Sub
End Class