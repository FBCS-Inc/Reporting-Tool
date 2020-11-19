Imports System.Data.SqlClient
Imports System.IO
Public Class SQLExport
    Dim Tmilli As Integer
    Dim Tsec As Integer
    Dim Tmin As Integer
    Dim Thour As Integer
    Dim dt1 As New DataTable
    Public RPTNAME As String
    Public shfmt As String
    Public shtlim As Integer
    Public vsql As String
    Public ext As String
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
    Private Sub SQLExport_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RptDataGridView.DataSource = dt1
        RPTNAME = "Query Result"
        ext = ".xlsx"
        FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + RPTNAME + ext
    End Sub

    Private Sub QueryButton_Click(sender As Object, e As EventArgs) Handles QueryButton.Click
        CheckForIllegalCrossThreadCalls = False
		vsql = SQLBox.Text.Replace("@customer", CustList1.Clist)

		ActivityTextBox.Text = "Running Query."
        WriteTime()
        StopWatch.Enabled = True
        QueryThread = New System.Threading.Thread(AddressOf RunQuery
)
        QueryThread.Start()
    End Sub
    Public Sub RunQuery()

        Try
            Cursor = Cursors.WaitCursor
            Dim cmd As New SqlCommand(vsql)
            Dim cn As New SqlConnection With {
                .ConnectionString = My.Settings.collect2000ConnectionString
            }
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(vsql, cn.ConnectionString)
			da1.SelectCommand.CommandTimeout = 1800
			'If TextBox1.Text <> "" Then
			'	da1.SelectCommand.Parameters.AddWithValue("@customer", TextBox1.Text)
			'End If
			dt1.Clear()
				ExportButton.Enabled = False
            cn.Open()
            da1.Fill(dt1)
            da1.Dispose()
            cn.Close()
            StopWatch.Stop()
            StopWatch.Enabled = False
            WriteTime()
            Tmilli = 0
            Tsec = 0
            Tmin = 0
            Thour = 0

            If dt1.Rows.Count > 0 Then
                ActivityTextBox.Text = "Returned " + Format(dt1.Rows.Count, "###,###,###,##0") + " Rows."
                ExportButton.Enabled = True
            Else
                ActivityTextBox.Text = "Returned nothing"
                ExportButton.Enabled = False
            End If

        Catch ex As System.Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
        End Try
        Cursor = Cursors.Default

    End Sub
    Private Sub Export()
        Try
            Cursor = Cursors.WaitCursor
            WriteTime()
            StopWatch.Enabled = True
            ' Dim ex = New ExportData.ExportData
            If ex Is Nothing Then
                ex = New ExportData With {
                .FirstRow = 2,
                .Visable = False,
                .ExcelFile = FileTextBox.Text,
                .SheetFormating = shfmt,
                .ActiveSheet = 1,
                .RowsPerSheet = shtlim
            }
            Else
                ex.ActiveSheet = 1
                ex.SheetFormating = shfmt
            End If


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
    Private Sub Excelfiletype_CheckedChanged(sender As Object, e As EventArgs) Handles excelfiletype.CheckedChanged
        If excelfiletype.Checked Then
            GroupBox4.Visible = excelfiletype.Checked
            ext = ".xlsx"
        Else
            GroupBox4.Visible = excelfiletype.Checked
            ext = ".csv"
        End If
    End Sub

    Private Sub ExportButton_Click(sender As Object, e As EventArgs) Handles ExportButton.Click
        CheckForIllegalCrossThreadCalls = False
        shfmt = Fmtbox.Text.ToUpper
        If SheetLimit.Checked Then
            shtlim = CInt(LimitBox.Text)
        Else
            shtlim = 1000000
        End If
        ActivityTextBox.Text = "Exporting " + Format(RptDataGridView.RowCount, "###,###,###,##0") + " Rows."
        WriteTime()
        StopWatch.Enabled = True
        ExportThread = New System.Threading.Thread(AddressOf Export)
        ExportThread.Start()
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

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            ext = ".csv"
        Else
            ext = ".xlsx"
        End If
    End Sub

    Private Sub Sheetall_CheckedChanged(sender As Object, e As EventArgs) Handles sheetall.CheckedChanged
        If sheetall.Checked Then
            shtlim = 1000000
        Else
            shtlim = CInt(LimitBox.Text)
        End If
    End Sub


End Class