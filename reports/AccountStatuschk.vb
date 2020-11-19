Imports System.Data.SqlClient

Public Class AccountStatuschk
	Private vsql As String = ""
	Private dtbl As New DataTable
	Private acct As String = ""
	Dim da1 As SqlDataAdapter
	Dim usql As String

	Private Sub Lookupbtn_Click(sender As Object, e As EventArgs) Handles Lookupbtn.Click
		StatCalc.Text = ""
		dtbl.Clear()
		Lookup()
	End Sub
	Private Sub Lookup()
		vsql = "select m.number,m.status,me.thedata,m.closed,m.returned,m.customer,n.created,n.action,n.result,n.comment "
		vsql += "from master m(nolock) "
		vsql += "inner join notes n(nolock)on n.number=m.number "
		vsql += "inner join miscextra me(nolock)on me.number=m.number and title='LastARCStatus' "
		vsql += "inner join fact(nolock)on fact.customerid=m.customer "
		vsql += "inner join customcustgroups ccg(nolock)on ccg.id=fact.CustomGroupID "
		vsql += "where m.account=@acct "
		vsql += "and ccg.Name like 'FIN/Absolute Resolutions' "
		vsql += "And n.created between @start And @End "
		If statusnotesbtn.Checked Then
			vsql += "And n.result In('RD','MG','D','AD','CV','CW')"
		End If
		vsql += "order by n.created desc"

		Try
			Cursor = Cursors.WaitCursor

			Dim cn As New SqlConnection With {
				.ConnectionString = My.Settings.collect2000ConnectionString
			}
			da1 = New SqlDataAdapter(vsql, cn.ConnectionString)
			da1.SelectCommand.CommandTimeout = 1800
			da1.SelectCommand.Parameters.AddWithValue("@acct", AccountNumber.Text)
			da1.SelectCommand.Parameters.AddWithValue("@start", StartDate.Text)
			da1.SelectCommand.Parameters.AddWithValue("@end", enddate.Text)

			cn.Open()
			da1.Fill(dtbl)
			da1.Dispose()
			cn.Close()

		Catch ex As Exception
			Cursor = Cursors.Default
			MessageBox.Show(ex.Message)
			Dim out As String = da1.SelectCommand.ToString()
		End Try
		Cursor = Cursors.Default
		FindStatus()
	End Sub
	Private Sub FindStatus()
		StatCalc.Text = ""
		For Each row As DataGridViewRow In DataGridView1.Rows
			row.Selected = True
			StatusCalc(row.Cells.Item(7).Value, row.Cells.Item(8).Value, row.Cells.Item(9).Value)
			If StatCalc.Text <> "" Then Exit For
		Next
	End Sub
	Private Sub AccountStatuschk_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		DataGridView1.DataSource = dtbl
	End Sub
	Private Sub StatusCalc(ByVal actn As String, ByVal rslt As String, ByVal comm As String)
		If IsNothing(comm) Then comm = ""
		If rslt = "RD" Then
			StatCalc.Text = "320221"
		ElseIf actn = "RV" And rslt = "MG" And comm.ToLower.Contains("recvd complaint verbal") Then
			StatCalc.Text = "112920"
		ElseIf actn = "RV" And rslt = "MG" And comm.ToLower.Contains("recvd complaint written") Then
			StatCalc.Text = "320000"
		ElseIf rslt = "D" And comm.ToLower.ToLower.Contains("oscar") Then
			StatCalc.Text = "320300"
		ElseIf rslt = "D" Then
			StatCalc.Text = "112910"
		ElseIf actn = "RV" And rslt = "AD" And comm.ToLower.Contains("written dispute") Then
			StatCalc.Text = "320220"
		ElseIf actn = "RV" And rslt = "AD" And comm.ToLower.Contains("summons recvd") Then
			StatCalc.Text = "320100"
		ElseIf actn = "RV" And rslt = "AD" And comm.ToLower.Contains("threat of suit") Then
			StatCalc.Text = "320200"
		ElseIf actn = "RV" And rslt = "AD" And comm.ToLower.Contains("hardship") Then
			StatCalc.Text = "320260"
		ElseIf actn = "RV" And rslt = "AD" And comm.ToLower.Contains("other 320300") Then
			StatCalc.Text = "320300"
		Else
			StatCalc.Text = ""
		End If
	End Sub


	Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
		DataGridView1.CurrentRow.Selected = True
		StatusCalc(DataGridView1.CurrentRow.Cells.Item(7).Value, DataGridView1.CurrentRow.Cells.Item(8).Value, DataGridView1.CurrentRow.Cells.Item(9).Value)
	End Sub

	Private Sub AccountNumber_TextChanged(sender As Object, e As EventArgs) Handles AccountNumber.TextChanged
		StatCalc.Text = ""
		dtbl.Clear()
	End Sub

	Private Sub StartDate_ValueChanged(sender As Object, e As EventArgs) Handles StartDate.ValueChanged
		StatCalc.Text = ""
		dtbl.Clear()

	End Sub

	Private Sub enddate_ValueChanged(sender As Object, e As EventArgs) Handles enddate.ValueChanged
		StatCalc.Text = ""
		dtbl.Clear()

	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		CheckForIllegalCrossThreadCalls = False

		UpdtData(dtbl)
	End Sub
	Public Sub UpdtData(ByRef dtbl As DataTable)
		usql = "update MiscExtra "
		usql += " set thedata=@newdata "
		usql += " where Number=@filenum "
		usql += " and Title='LastARCStatus' and TheData = @olddata "
		Try
			Cursor = Cursors.WaitCursor
			Dim cmd As New SqlCommand(usql)
			Dim cn As New SqlConnection With {
				.ConnectionString = My.Settings.collect2000ConnectionString
			}
			Dim da1 As SqlCommand = New SqlCommand(usql, cn) With {
				.CommandTimeout = 1800
			}

			da1.Parameters.AddWithValue("@filenum", DataGridView1.CurrentRow.Cells(0).Value)
			da1.Parameters.AddWithValue("@olddata", DataGridView1.CurrentRow.Cells(2).Value)
			da1.Parameters.AddWithValue("@newdata", StatCalc.Text)
			cn.Open()
			da1.ExecuteNonQuery()
			da1.Dispose()
			StatCalc.Text = ""
			dtbl.Clear()
			Lookup()

		Catch ex As Exception
			Cursor = Cursors.Default
			MessageBox.Show(ex.Message)
		End Try
		Cursor = Cursors.Default

	End Sub
End Class