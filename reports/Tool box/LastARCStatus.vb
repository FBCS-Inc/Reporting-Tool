Imports System.Data.SqlClient

Public Class LastARCStatus
    Dim vsql As String
    Dim usql As String
    Dim dt1 As New DataTable
    Dim dt2 As New DataTable
    Private UpdateThread As System.Threading.Thread
    Private QueryThread As System.Threading.Thread



    Private Sub LastARCStatus_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BeforeDataGridView.DataSource = dt1
        AfterDataGridView.DataSource = dt2
        vsql = "select m.number,m.account,m.status,isnull(me1.thedata,'') [Last ARC Status], "
        vsql += " isnull(me2.TheData,'') [Status Date],ISNULL(me3.thedata,'') [Preissue Status] "
        vsql += " From master m with(nolock) "
        vsql += " inner join MiscExtra me1 with(nolock)on me1.Number=m.number and me1.title='LastARCStatus'"
        vsql += " Left join MiscExtra me2 with(nolock)on me2.Number=m.number and me2.title='LastARCStatusDate'"
        vsql += " Left join MiscExtra me3 with(nolock)on me3.Number=m.number and me3.title = 'PreIssueARCStatus'"
        vsql += " where m.account = @account"

        usql = "update MiscExtra "
        usql += " set thedata=@newdata "
        usql += " where Number=@filenum "
        usql += " and Title='LastARCStatus' and TheData = @olddata "

    End Sub

    Public Sub ViewData(ByRef dtbl As DataTable)

        Try
            Cursor = Cursors.WaitCursor
			'Dim cmd As New SqlCommand(vsql)
			Dim cn As New SqlConnection With {
                .ConnectionString = My.Settings.collect2000ConnectionString
            }
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(vsql, cn.ConnectionString)
            da1.SelectCommand.CommandTimeout = 1800

            da1.SelectCommand.Parameters.AddWithValue("@account", Accountbox.Text)

            dtbl.Clear()
            AfterButton.Enabled = False
            cn.Open()
            da1.Fill(dtbl)
            da1.Dispose()
            cn.Close()


            If dtbl.Rows.Count > 0 Then
                CurrentARC.Text = dtbl.Rows(0).Item(3).ToString
                Filebox.Text = dtbl.Rows(0).Item(0).ToString
                DateBox.Text = dtbl.Rows(0).Item(4).ToString
                prebox.Text = dtbl.Rows(0).Item(5).ToString

            Else
                CurrentARC.Text = ""
                Filebox.Text = ""
                DateBox.Text = ""
                prebox.Text = ""

            End If

        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try
        Cursor = Cursors.Default

    End Sub
    Public Sub UpdtData(ByRef dtbl As DataTable)

        Try
            Cursor = Cursors.WaitCursor
            Dim cmd As New SqlCommand(usql)
            Dim cn As New SqlConnection With {
                .ConnectionString = My.Settings.collect2000ConnectionString
            }
            Dim da1 As SqlCommand = New SqlCommand(usql, cn) With {
                .CommandTimeout = 1800
            }

            da1.Parameters.AddWithValue("@filenum", Filebox.Text)
            da1.Parameters.AddWithValue("@olddata", CurrentARC.Text)
			da1.Parameters.AddWithValue("@newdata", NewARC.Text)
			cn.Open()
			da1.ExecuteNonQuery()
            da1.Dispose()
            AfterButton.Enabled = False
            ViewData(dt2)

        Catch ex As Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try
        Cursor = Cursors.Default

    End Sub
    Private Sub BeforeButton_Click(sender As Object, e As EventArgs) Handles BeforeButton.Click
        CheckForIllegalCrossThreadCalls = False

        ViewData(dt1)
    End Sub
	Private Sub Reset()
		NewARC.Text = ""
		CurrentARC.Text = ""
		Filebox.Text = ""
		DateBox.Text = ""
		prebox.Text = ""
		dt1.Clear()
		dt2.Clear()
	End Sub
	Private Sub Accountbox_TextChanged(sender As Object, e As EventArgs) Handles Accountbox.TextChanged
		BeforeButton.Enabled = True
		Reset()
	End Sub

    Private Sub NewARC_TextChanged(sender As Object, e As EventArgs) Handles NewARC.TextChanged
        If Len(NewARC.Text) = 6 Then
            AfterButton.Enabled = True
        End If
    End Sub

	Private Sub AfterButton_Click(sender As Object, e As EventArgs) Handles AfterButton.Click
		CheckForIllegalCrossThreadCalls = False

		UpdtData(dt1)
	End Sub
End Class