Imports System.Data.SqlClient

Public Class ARCStatusHistory
	Private vsql As String = ""
	Private dtbl As New DataTable
	Private acct As String = ""
	Dim da1 As SqlDataAdapter
	Private Sub ARCStatusHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		DataGridView1.DataSource = dtbl
	End Sub

	Private Sub Lookupbtn_Click(sender As Object, e As EventArgs) Handles Lookupbtn.Click
		acct = "'" + TextBox1.Text + "'"
		vsql = "Select Account,DateChanged,ARCStatus"
		vsql += " FROM ARCStatusUpdateHistory"
		vsql += " where Account = @acct "
		If preissue.Checked = True Then
			vsql += " and ARCStatus not in ('320270','320250','320221','112920','320000','112910','320220','320250','320100','320200','320260','320300')"
		End If
		vsql += " order by DateChanged desc"
		Try
			Cursor = Cursors.WaitCursor

			Dim cn As New SqlConnection With {
				.ConnectionString = My.Settings.collect2000ConnectionString
			}
			da1 = New SqlDataAdapter(vsql, cn.ConnectionString)
			da1.SelectCommand.CommandTimeout = 1800
			da1.SelectCommand.Parameters.AddWithValue("@acct", TextBox1.Text)

			dtbl.Clear()
			cn.Open()
			da1.Fill(dtbl)
			If dtbl.Rows.Count < 1 Then MessageBox.Show("No History returned for this account")
			da1.Dispose()
			cn.Close()

		Catch ex As Exception
			Cursor = Cursors.Default
			MessageBox.Show(ex.Message)
			Dim out As String = da1.SelectCommand.ToString()
		End Try
		Cursor = Cursors.Default
	End Sub
End Class