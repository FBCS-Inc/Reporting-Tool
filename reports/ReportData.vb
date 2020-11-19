Imports System.Data.SqlClient
Imports System.IO

Public Class ReportData
    Dim Tmilli As Integer
    Dim Tsec As Integer
    Dim Tmin As Integer
    Dim Thour As Integer
    Dim dt1 As New DataTable
    Public RPTNAME As String
    Public shfmt As String
    Public vsql As String
    Public ext As String
	Public rptdte As String = "_" + Format(Today, "yyyyMMdd").ToString
	Private ExportThread As System.Threading.Thread
    Private QueryThread As System.Threading.Thread
    Private ex As ExportData = Nothing

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
    Private Sub ReportData_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StartDate.Value = DateAdd(DateInterval.Day, -1, Today)
        enddate.Value = DateAdd(DateInterval.Day, -1, Today)
        ext = ".xlsx"
        RadioButton2.Checked = True
        RptDataGridView.DataSource = dt1

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        FileTextBox.Text = OpenFileDialog1.FileName

    End Sub

    Private Sub FileButton_Click(sender As Object, e As EventArgs) Handles FileButton.Click
        With OpenFileDialog1
            .FileName = FileTextBox.Text
            .Filter = "CSV files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx"
            .ShowDialog()
        End With

    End Sub

    Private Sub GenerateButton_Click(sender As Object, e As EventArgs) Handles GenerateButton.Click
        CheckForIllegalCrossThreadCalls = False

        ActivityTextBox.Text = "Running Report."
        WriteTime()
        StopWatch.Enabled = True
        QueryThread = New System.Threading.Thread(AddressOf GetReport)
        QueryThread.Start()
    End Sub

	Private Sub ExportButton_Click(sender As Object, e As EventArgs) Handles ExportButton.Click
		CheckForIllegalCrossThreadCalls = False

		ActivityTextBox.Text = "Exporting " + Format(RptDataGridView.RowCount, "###,###,###,##0") + " Rows."
		WriteTime()
		StopWatch.Enabled = True
		ExportThread = New System.Threading.Thread(AddressOf Export)
		ExportThread.Start()
	End Sub
	Private Sub Exportsetup()
		ex = New ExportData With {
				.FirstRow = 2,
				.Visable = False,
				.ExcelFile = FileTextBox.Text,
				.SheetFormating = shfmt,
				.ActiveSheet = 1
				 }
	End Sub
	Private Sub Export()
        Try
            Cursor = Cursors.WaitCursor
            WriteTime()
            StopWatch.Enabled = True

			If ex Is Nothing Then Exportsetup()
			'ex = New ExportData With {
			'             .FirstRow = 2,
			'             .Visable = False,
			'             .ExcelFile = FileTextBox.Text,
			'             .SheetFormating = shfmt,
			'             .ActiveSheet = 1
			'              }
			'Else
			'    ex.ActiveSheet = 1
			'    ex.SheetFormating = shfmt
			'End If


			ActivityTextBox.Text = "Writing " + Format(dt1.Rows.Count, "###,###,###,##0") + " Rows to the report."
            ex.Data = dt1
            ex.Export()
            Cursor = Cursors.Default
            MessageBox.Show("Report: " + FileTextBox.Text + " is complete", "Report Complete")

        Catch ex As System.Exception

            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        Finally
            StopWatch.Stop()
            StopWatch.Enabled = False
            WriteTime()
            Tmilli = 0
            Tsec = 0
            Tmin = 0
            Thour = 0
        End Try
        Cursor = Cursors.Default
    End Sub

    Public Sub GetReport()

        Try
            Cursor = Cursors.WaitCursor
            Dim cmd As New SqlCommand(vsql)
            Dim cn As New SqlConnection With {
                .ConnectionString = My.Settings.collect2000ConnectionString
            }
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(vsql, cn.ConnectionString)
            da1.SelectCommand.CommandTimeout = 1800
            If StartDate.Visible Then
                da1.SelectCommand.Parameters.AddWithValue("@start", StartDate.Text)
            End If
            If enddate.Visible Then
                da1.SelectCommand.Parameters.AddWithValue("@end", enddate.Text)
            End If
            dt1.Clear()
            ExportButton.Enabled = False
            cn.Open()
			da1.Fill(dt1)

UserAborted: da1.Dispose()
			cn.Close()
            StopWatch.Stop()
            StopWatch.Enabled = False
            WriteTime()
            Tmilli = 0
            Tsec = 0
            Tmin = 0
            Thour = 0

            If dt1.Rows.Count > 0 Then
                'RptDataGridView.
                'RptDataGridView.DataSource = dt1
                ActivityTextBox.Text = "Returned " + Format(dt1.Rows.Count, "###,###,###,##0") + " Rows."
                ExportButton.Enabled = True
            Else
                'RptDataGridView.DataSource = dt1
                ActivityTextBox.Text = "Returned nothing"
                ExportButton.Enabled = False
            End If

        Catch ex As System.Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try
        Cursor = Cursors.Default

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        ext = ".csv"
        FileTextBox.Text = Path.ChangeExtension(FileTextBox.Text, ext)
		Exportsetup()
	End Sub

	Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
		ext = ".xlsx"
		FileTextBox.Text = Path.ChangeExtension(FileTextBox.Text, ext)
		Exportsetup()
	End Sub
	Public Sub StopQuery()
		QueryThread.Abort()
		ActivityTextBox.Text = "Query Aborted by user."
		StopWatch.Stop()
		StopWatch.Enabled = False
		WriteTime()
		Tmilli = 0
		Tsec = 0
		Tmin = 0
		Thour = 0

	End Sub
	Public Sub StopExport()
		ExportThread.Abort()
		ActivityTextBox.Text = "Export Aborted by user."
		StopWatch.Stop()
		StopWatch.Enabled = False
		WriteTime()
		Tmilli = 0
		Tsec = 0
		Tmin = 0
		Thour = 0
	End Sub

	Private Sub FileTextBox_TextChanged(sender As Object, e As EventArgs) Handles FileTextBox.TextChanged
		Exportsetup()
	End Sub
End Class
