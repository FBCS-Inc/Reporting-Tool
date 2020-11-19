Imports System.Data.SqlClient

Public Class DateMaintenance
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		GetFileData()

	End Sub
	Private Sub GetFileData()
		Dim dt1 As New DataTable
		Dim vsql As String
		vsql = "select m.number,m.closed,m.Returned "
		vsql += "From master m with(nolock) "
		vsql += "where m.number = '" + filenumber.Text.Trim + "'"
		vsql += "and m.closed is not null"
		Cursor = Cursors.WaitCursor
		Dim cmd As New SqlCommand(vsql)
		Dim cn As New SqlConnection With {
			.ConnectionString = My.Settings.collect2000ConnectionString
		}
		Dim da1 As SqlDataAdapter = New SqlDataAdapter(vsql, cn.ConnectionString)
		da1.SelectCommand.CommandTimeout = 1800
		dt1.Clear()
		cn.Open()
		da1.Fill(dt1)
		da1.Dispose()
		cn.Close()
		If dt1.Rows.Count = 1 Then
			Enableform()
			closedbox.Text = dt1.Rows(0).Item(1)
			If IsDBNull(dt1.Rows(0).Item(2)) Then
				returnedbox.Text = ""
			Else
				returnedbox.Text = dt1.Rows(0).Item(2)
			End If
		ElseIf dt1.Rows.Count > 1 Then
			Disableform()
		Else
			Disableform()
			MessageBox.Show("This account is not closed and cant be adjusted")
		End If

		Cursor = Cursors.Default
	End Sub
	Private Sub Enableform()
		Me.ClosedDate.Enabled = True
		Me.ReturnedDate.Enabled = True
		Me.Button2.Enabled = True
		Me.Button3.Enabled = True
	End Sub
	Private Sub Disableform()
		Me.ClosedDate.Enabled = False
		Me.ReturnedDate.Enabled = False
		Me.Button2.Enabled = False
		Me.Button3.Enabled = False
	End Sub
	Private Sub DateMaintenance_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		ClosedDate.MaxDate = Now
		ReturnedDate.MaxDate = Now
		Disableform()
	End Sub


	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		If Validatedate(ClosedDate.Text) Then

			Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
			Dim msg = "You are about to change the Closed date for this File Number. 
Do you want to continue?"
			Dim title = "Data Changing Request"
			Dim response = MsgBox(msg, style, title)
			If response = MsgBoxResult.Yes Then
				msg = "This will change the closed date to: " + ClosedDate.Value.ToShortDateString + " 
Are you certain this is what you want to do?"
				response = MsgBox(msg, style, title)
				If response = MsgBoxResult.Yes Then SaveClosed()
			End If
		Else
			MessageBox.Show("Please enter a valid date for the closed date")
		End If
	End Sub
	Private Sub SaveClosed()
		Dim vsql As String
		vsql = "update master "
		vsql += "set closed = '" + ClosedDate.Value.ToShortDateString.ToString + "'"
		vsql += "where number = '" + filenumber.Text.Trim + "'"
		vsql += "insert into notes"
		vsql += "([number],[created],[user0],[action],[result],[comment])"
		vsql += "values('" + filenumber.Text.Trim + "',GETDATE(),'dtemnt','Date','Chng','Closed date changed from " + closedbox.Text + ", To " + ClosedDate.Value.ToShortDateString.ToString + "')"
		Cursor = Cursors.WaitCursor

		Dim cn As New SqlConnection With {
			.ConnectionString = My.Settings.collect2000ConnectionString
		}
		cn.Open()
		Dim cmd As New SqlCommand(vsql, cn)
		cmd.ExecuteNonQuery()
		cn.Close()
		GetFileData()
	End Sub
	Private Sub SaveReturned()
		Dim vsql As String
		vsql = "update master "
		vsql += "set returned = '" + ReturnedDate.Value.ToShortDateString.ToString + "'"
		vsql += "where number = '" + filenumber.Text.Trim + "'"
		vsql += "insert into notes"
		vsql += "([number],[created],[user0],[action],[result],[comment])"
		vsql += "values('" + filenumber.Text.Trim + "',GETDATE(),'dtemnt','Date','Chng','Returned date changed from " + returnedbox.Text + ", To " + ReturnedDate.Value.ToShortDateString.ToString + "')"

		Cursor = Cursors.WaitCursor

		Dim cn As New SqlConnection With {
			.ConnectionString = My.Settings.collect2000ConnectionString
		}
		cn.Open()
		Dim cmd As New SqlCommand(vsql, cn)
		cmd.ExecuteNonQuery()
		cn.Close()
		GetFileData()
	End Sub
	Private Function Validatedate(dte As String) As Boolean

		If Not IsDate(dte) Then Return False

		Dim dtest As DateTime = Convert.ToDateTime(dte)


		If dtest > Now Then Return False
		Return True
	End Function
	Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
		If Validatedate(ReturnedDate.Text) Then

			Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
			Dim msg = "You are about to change the Returned date for this File Number. 
Do you want to continue?"
			Dim title = "Data Changing Request"
			Dim response = MsgBox(msg, style, title)
			If response = MsgBoxResult.Yes Then
				msg = "This will change the Returned date to: " + ReturnedDate.Value.ToShortDateString + " 
Are you certain this is what you want to do?"
				response = MsgBox(msg, style, title)
				If response = MsgBoxResult.Yes Then SaveReturned()
			End If
		Else
			MessageBox.Show("Please enter a valid date for the returned date")
		End If
	End Sub

	Private Sub filenumber_TextChanged(sender As Object, e As EventArgs) Handles filenumber.TextChanged
		Disableform()
	End Sub
End Class