Imports System.Windows.Forms

Public Class Reporting

	Private Sub Reports_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        IsMdiContainer = True
		If Environment.UserName = "edward.hall" Then ToolsToolStripMenuItem.Visible = True
		If Environment.UserName = "mary.albert" Then ToolsToolStripMenuItem.Visible = True
		Me.Text = "Reporting Tool --- Version(" + Application.ProductVersion + ")"
	End Sub
	Private Sub PPACreditAuditToolStripMenuItem1_Click(sender As Object, e As EventArgs)

		OpenChild(My.Forms.PPACreditAudit)
	End Sub
	Private Sub SettlementsToBe2ndReviewedToolStripMenuItem1_Click(sender As Object, e As EventArgs)

		OpenChild(My.Forms.Settlement)
	End Sub
	Private Sub OpenChild(cf As Form)
		CloseChildren()

		cf.MdiParent = Me
        cf.Dock = DockStyle.Fill
		cf.WindowState = FormWindowState.Maximized
		cf.ControlBox = False
		cf.ShowIcon = False
		cf.FormBorderStyle = FormBorderStyle.None
		cf.HelpButton = False
		cf.Icon = Nothing
		cf.MaximizeBox = False
		cf.MinimizeBox = False

		cf.Show()
    End Sub

    Private Sub CloseChildren()
        For Each cf As Form In Me.MdiChildren
            cf.Close()
        Next
    End Sub
    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        CloseChildren()
    End Sub

    Private Sub ExeterSifPifToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExeterSifPifToolStripMenuItem.Click
        OpenChild(My.Forms.ExeterSifPif)
    End Sub

    Private Sub TimezonesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TimezonesToolStripMenuItem.Click
        OpenChild(My.Forms.Exception8)
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem2.Click
        OpenChild(My.Forms.Exception2)
    End Sub

    Private Sub Inventory3160DaysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Inventory3160DaysToolStripMenuItem.Click
        OpenChild(My.Forms.Exception5)
    End Sub

    Private Sub Inventory6190DaysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Inventory6190DaysToolStripMenuItem.Click
        OpenChild(My.Forms.Exception6)
    End Sub

    Private Sub First30DaysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles First30DaysToolStripMenuItem.Click
        OpenChild(My.Forms.Exception4)
    End Sub

    Private Sub InventoryWorkGapToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InventoryWorkGapToolStripMenuItem.Click
        OpenChild(My.Forms.Exception3)
    End Sub

    Private Sub ContactsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactsToolStripMenuItem.Click
        OpenChild(My.Forms.Exception7)
    End Sub

    Private Sub DialerContactedInLessThan8DaysToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DialerContactedInLessThan8DaysToolStripMenuItem.Click
        OpenChild(My.Forms.Exception9)
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        OpenChild(My.Forms.AboutBox1)
    End Sub

    Private Sub IinstructionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IinstructionsToolStripMenuItem.Click
        OpenChild(My.Forms.Help)
    End Sub



    Private Sub LastARCStatusToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LastARCStatusToolStripMenuItem.Click
        OpenChild(My.Forms.LastARCStatus)
    End Sub

    Private Sub DebtSettlementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DebtSettlementToolStripMenuItem.Click
        OpenChild(My.Forms.DebtSettlement)
    End Sub

    Private Sub GenericQueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenericQueryToolStripMenuItem.Click
        OpenChild(My.Forms.SQLExport)
    End Sub

    Private Sub ExitToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub AccountQueryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountQueryToolStripMenuItem.Click
        OpenChild(My.Forms.AccountQuery)
    End Sub

    Private Sub PendingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PendingToolStripMenuItem.Click
        OpenChild(My.Forms.PendingLetterLinkExp)
    End Sub

    Private Sub ProcessedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProcessedToolStripMenuItem.Click
        OpenChild(My.Forms.ProcessedLetterLinkExp)
    End Sub

    Private Sub PDCDeletedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PDCDeletedToolStripMenuItem.Click
        OpenChild(My.Forms.PCDDeletedReport)
    End Sub

    Private Sub AFFInputFixToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AFFInputFixToolStripMenuItem.Click
        OpenChild(My.Forms.AFFfix)
    End Sub

	Private Sub TSISMERestrictionRemovalToolStripMenuItem_Click(sender As Object, e As EventArgs) 
		OpenChild(My.Forms.TSISMERestrictionRemoval)
	End Sub

	Private Sub AFFFileCleanUpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AFFFileCleanUpToolStripMenuItem.Click
		OpenChild(My.Forms.AFFfix)
	End Sub

	Private Sub DateMaintenanceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DateMaintenanceToolStripMenuItem.Click
		OpenChild(My.Forms.DateMaintenance)
	End Sub

	Private Sub CountToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CountToolStripMenuItem.Click
		OpenChild(My.Forms.autobuildercount)
	End Sub

	Private Sub Count2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Count2ToolStripMenuItem.Click
		OpenChild(My.Forms.autobuildercount2)
	End Sub

	Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
		OpenChild(My.Forms.NewAgent)
	End Sub

	Private Sub Recording2WeekSummaryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Recording2WeekSummaryToolStripMenuItem.Click
		OpenChild(My.Forms.recordings2weeks)
	End Sub

	Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
		OpenChild(My.Forms.Recordingcount)
	End Sub

	Private Sub ExeterBankruptcyClosesToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.ExeterBankruptcyCloses)
	End Sub

	Private Sub ArrangementsGreaterThan45DaysToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.Arrang45days)
	End Sub
	Private Sub PPAExceptionToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.PPAException)
	End Sub
	Private Sub AbortToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbortToolStripMenuItem.Click
		Dim rd1 As ReportData = ActiveMdiChild.Controls().Item("ReportData1")
		rd1.StopQuery()
		rd1.StopExport()
	End Sub
	Private Sub PPACreaditExceptionToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.PPACredit1)
	End Sub
	Private Sub PPAAuditToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.PPAAudit)
	End Sub
	Private Sub ToolStripMenuItem6_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
		OpenChild(My.Forms.PPAAudit)
	End Sub

	Private Sub ToolStripMenuItem7_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem7.Click
		OpenChild(My.Forms.PPAException)
	End Sub

	Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem8.Click
		OpenChild(My.Forms.PPACredit1)
	End Sub

	Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem9.Click
		OpenChild(My.Forms.PPACreditAudit)
	End Sub

	Private Sub ToolStripMenuItem10_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem10.Click
		OpenChild(My.Forms.Removed737Q)
	End Sub

	Private Sub ToolStripMenuItem11_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem11.Click
		OpenChild(My.Forms.Arrang45days)
	End Sub

	Private Sub ToolStripMenuItem12_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem12.Click
		OpenChild(My.Forms.ExeterBankruptcyCloses)
	End Sub

	Private Sub SettlementsToBe2ndReviewedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettlementsToBe2ndReviewedToolStripMenuItem.Click
		OpenChild(My.Forms.Settlement)
	End Sub

	Private Sub ChangeCustomerNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeCustomerNumberToolStripMenuItem.Click
		OpenChild(My.Forms.ChangeClientNumber)
	End Sub

	Private Sub FeeCodeChangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FeeCodeChangeToolStripMenuItem.Click
		OpenChild(My.Forms.FeeCodeChg)
	End Sub

	Private Sub CustomerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CustomerToolStripMenuItem.Click
		OpenChild(My.Forms.SelectCustomer)
	End Sub

	Private Sub ClosedReturnedDatesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClosedReturnedDatesToolStripMenuItem.Click
		OpenChild(My.Forms.DateMaintenance)
	End Sub

	'Private Sub UstoimerNumberChangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UstoimerNumberChangeToolStripMenuItem.Click
	'	OpenChild(My.Forms.ChangeClientNumber)
	'End Sub

	Private Sub ARCStatusHistoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ARCStatusHistoryToolStripMenuItem.Click
		OpenChild(My.Forms.ARCStatusHistory)
	End Sub

	Private Sub ReturnedDateMassChangeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReturnedDateMassChangeToolStripMenuItem.Click
		OpenChild(My.Forms.ChangeReturnDate)
	End Sub

	Private Sub AccountStatusCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AccountStatusCheckToolStripMenuItem.Click
		OpenChild(My.Forms.AccountStatuschk)
	End Sub

	Private Sub EmailUpdateToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EmailUpdateToolStripMenuItem.Click
		OpenChild(My.Forms.Emailupdate)
	End Sub

	Private Sub TestToolStripMenuItem_Click(sender As Object, e As EventArgs)
		OpenChild(My.Forms.Customer)
	End Sub

	Private Sub PostDatedTransDeletedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PostDatedTransDeletedToolStripMenuItem.Click
		OpenChild(My.Forms.PostDatedTransDeleted)
	End Sub

	Private Sub MassChangeReturnedDatesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MassChangeReturnedDatesToolStripMenuItem.Click
		OpenChild(My.Forms.ChangeReturnDate)
	End Sub

	Private Sub ChangeReceivedDatesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeReceivedDatesToolStripMenuItem.Click
		OpenChild(My.Forms.Change_Receive_date)
	End Sub

	Private Sub ChangeAccountNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChangeAccountNumberToolStripMenuItem.Click
		OpenChild(My.Forms.ChangeAccountNumber)
	End Sub

	Private Sub ContactLetterExceptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ContactLetterExceptionsToolStripMenuItem.Click
		OpenChild(My.Forms.ContactLetterExceptions)
	End Sub
End Class
