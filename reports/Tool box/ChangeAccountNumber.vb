Imports System.Data.SqlClient
Imports System.IO
Imports Microsoft.Office.Interop

Public Class ChangeAccountNumber
	Dim Tmilli As Integer
	Dim Tsec As Integer
	Dim Tmin As Integer
	Dim Thour As Integer
	Private LoadThread As System.Threading.Thread
	Private QueryThread As System.Threading.Thread
	Dim dts As DataSet
	Dim excel As String
	Private fn As String
	Dim dt As DataTable
	Dim fi As FileInfo
	Dim count As Integer
	Dim Tcount As Integer
	Dim Mcount As Integer
	Dim Rcount As Integer
	Dim col As Integer
	Dim fcol As Integer
	Dim ocol As Integer
	Dim lines() As String
	Private help As String
	Dim cn As SqlConnection
	Private vsql = "declare @oldCode as varchar(10)
 set @oldcode=(select m.account from master m with(nolock) where number=@number)

if @oldcode is not null begin

	INSERT INTO notes
		(number, created, user0, [action], result, comment)
		VALUES
		(@number, getdate(), 'SYSTEM', 'CUST', 'CHNG', 'ACCOUNT# CHANGED | ' + @oldcode + ' | ' + @newcode)

INSERT INTO miscextra
		(number, title, thedata)
		VALUES
		(@number, 'PriorAcct#',  @oldcode)

	update master
	set account = @newcode
	from master with (rowlock)
	where number = @number

end
"
	Private Sub Stopwatch_Tick(sender As Object, e As EventArgs) Handles StopWatch.Tick
		Tmilli += 1
		If Tmilli = 100 Then
			Tsec += 1
			Tmilli = 0
		ElseIf Tsec = 60 Then
			Tmin += 1
			Tsec = 0
		ElseIf Tmin = 60 Then
			Thour += 1
			Tmin = 0
		End If
		WriteTime()
	End Sub
	Private Sub WriteTime()
		ElapseTextBox.Text = "Elapsed Time " + Strings.Right("0" + Thour.ToString, 2) + " : " + Strings.Right("0" + Tmin.ToString, 2) + " : " + Strings.Right("0" + Tsec.ToString, 2)
	End Sub

	Private Sub Inputxlsx(ByVal filename As String)
		Dim excel As New Excel.Application With {
			.DisplayAlerts = False
		}
		Dim workbook As Excel.Workbook = excel.Workbooks.Open(filename)
		Dim sheet As Excel.Worksheet = workbook.Sheets(1)

		filename = Path.ChangeExtension(filename, ".txt")
		sheet.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVMSDOS)

		workbook.Close()
		workbook = Nothing

		excel.Quit()
		excel = Nothing
		CheckForIllegalCrossThreadCalls = False
		LoadThread = New System.Threading.Thread(AddressOf ReadFile)
		LoadThread.Start(filename)
		'ReadFile(filename)


	End Sub
	Private Sub ReadFile(ByVal fn As String)
		WriteTime()
		Stopwatch.Enabled = True
		ActivityTextBox.Text = " Loading Excel file"
		Dim lines() As String = IO.File.ReadAllLines(fn)

		' add columns only if DGV has none
		col = 0
		If DataGridView1.Columns.Count = 0 Then
			For Each cnam In Split(lines(0), ",")
				DataGridView1.Columns.Add(cnam, cnam)
				ComboBox1.Items.Add(cnam)
				ComboBox2.Items.Add(cnam)
				col += 1
			Next

		End If
		Tcount = 0
		For x As Integer = 1 To UBound(lines)
			DataGridView1.Rows.Add(Split(lines(x).Replace("""", ""), ","))
			Tcount += 1
			WriteTime()

		Next
		Cursor = Cursors.Default
		ACount.Text = Tcount
		'FileColumn.Maximum = col
		'OldColumn.Maximum = col
		Stopwatch.Stop()
		Stopwatch.Enabled = False
		WriteTime()
		Tmilli = 0
		Tsec = 0
		Tmin = 0
		Thour = 0
		ActivityTextBox.Text = "Successful excel load"
	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

		OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
		OpenFileDialog1.Filter = "All Files (*.*)|*.*|Excel files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|XLS Files (*.xls)|*xls"

		OpenFileDialog1.ShowDialog()
	End Sub

	Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
		Me.Cursor = Cursors.WaitCursor
		fn = OpenFileDialog1.FileName
		fi = New FileInfo(fn)
		excel = fi.FullName
		Inputxlsx(fn)
	End Sub

	Private Sub MakeChange_Click(sender As Object, e As EventArgs) Handles MakeChange.Click
		CheckForIllegalCrossThreadCalls = False
		Dim style = MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2 Or MsgBoxStyle.Critical
		Dim msg = "You are about to make a mass change to client numbers. 
Do you want to continue?"
		Dim title = "Data Changing Request"
		Dim response = MsgBox(msg, style, title)
		If response = MsgBoxResult.Yes Then
			msg = "This will change the Client numbers of these accounts. 
Are you certain this is what you want to do?"
			response = MsgBox(msg, style, title)
			If response = MsgBoxResult.Yes Then
				ActivityTextBox.Text = "Changing old customer code "
				WriteTime()
				Stopwatch.Enabled = True
				QueryThread = New System.Threading.Thread(AddressOf RunQuery
)
				QueryThread.Start()
			End If
		End If

	End Sub
	Public Sub RunQuery()
		WriteTime()
		Stopwatch.Enabled = True

		Try
			Cursor = Cursors.WaitCursor
			'Dim cmd As New SqlCommand(vsql)

			cn.Open()
			Dim cm1 As SqlCommand = New SqlCommand(vsql, cn)
			cm1.Parameters.AddWithValue("@number", "000000000")
			cm1.Parameters.AddWithValue("@newcode", "000000000")
			'	cm1.Parameters.AddWithValue("@newcode", TextBox1.Text)
			count = 0
			For Each row As DataGridViewRow In DataGridView1.Rows
				row.Selected = True
				DataGridView1.FirstDisplayedScrollingRowIndex = row.Index
				Dim obj(row.Cells.Count - 1) As Object

				cm1.Parameters("@number").Value = row.Cells(fcol).Value
				cm1.Parameters("@newcode").Value = row.Cells(ocol).Value
				Rcount = cm1.ExecuteNonQuery()
				If Rcount > 0 Then count += 1

				WriteTime()
				row.Selected = False
			Next
			cn.Close()
			Stopwatch.Stop()
			Stopwatch.Enabled = False
			WriteTime()
			Tmilli = 0
			Tsec = 0
			Tmin = 0
			Thour = 0
			ActivityTextBox.Text = "Changed " + count.ToString + " email address."

		Catch ex As System.Exception
			Cursor = Cursors.Default
			MessageBox.Show(ex.Message)
			ActivityTextBox.Text = "Error in Changing email address."
		End Try
		Cursor = Cursors.Default
		Mcount = Tcount - count
		If Mcount > 0 Then
			MessageBox.Show(Mcount.ToString + " Accounts not found to update.")
		End If
	End Sub
	Private Function Mkchgenable() As Boolean
		If fcol = ocol Then
			Return False
		ElseIf DataGridView1.Rows.Count = 0 Then
			Return False
		ElseIf prod.Checked = testdb.Checked Then
			Return False
		Else
			Return True
		End If
	End Function

	Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)
		MakeChange.Enabled = Mkchgenable()
	End Sub

	Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
		fcol = ComboBox1.SelectedIndex
		MakeChange.Enabled = Mkchgenable()
	End Sub

	Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
		ocol = ComboBox2.SelectedIndex
		MakeChange.Enabled = Mkchgenable()
	End Sub

	Private Sub Prod_CheckedChanged(sender As Object, e As EventArgs) Handles prod.CheckedChanged
		If prod.Checked Then
			cn = New SqlConnection With {
				.ConnectionString = My.Settings.collect2000ConnectionString
			}
		ElseIf testdb.Checked Then
			cn = New SqlConnection With {
				.ConnectionString = My.Settings.Testdb
			}
		End If
		MakeChange.Enabled = Mkchgenable()
	End Sub

	Private Sub Testdb_CheckedChanged(sender As Object, e As EventArgs) Handles testdb.CheckedChanged
		If prod.Checked Then
			cn = New SqlConnection With {
				.ConnectionString = My.Settings.collect2000ConnectionString
			}
		ElseIf testdb.Checked Then
			cn = New SqlConnection With {
				.ConnectionString = My.Settings.Testdb
			}
		End If
		MakeChange.Enabled = Mkchgenable()
	End Sub

	Private Sub ChangeClientNumber_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		help = "1.) Build a spreadsheet with file # and the new customer numbers to change them to."
		help += Chr(13)
		help += "2.) Load the spreadsheet with filenumbers and new customers seperared in columns."
		help += Chr(13)
		help += "3.) Once the spread sheet loads select the column with file numbers in the File # Column list"
		help += Chr(13)
		help += "4.) Select the column with new numbers in the new code list"
		help += Chr(13)
		help += "5.) Select Production or Test database you want to make the changes too."
		help += Chr(13)
		help += "6.) Press the Make changes button to start the change."
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		MessageBox.Show(help)
	End Sub
End Class
