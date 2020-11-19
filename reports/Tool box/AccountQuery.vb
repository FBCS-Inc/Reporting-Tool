Imports System.Data.SqlClient
Imports System.IO
Public Class AccountQuery

    Dim dt1 As New DataTable
    Public RPTNAME As String
    Public shfmt As String
    Public shtlim As Integer
    Public vsql As String
    Public ext As String
    Private ExportThread As System.Threading.Thread
    Private QueryThread As System.Threading.Thread
    Private ex As ExportData = Nothing


    Private Sub AccountQuery_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RptDataGridView.DataSource = dt1
        RPTNAME = "Account Query Result"
        ext = ".xlsx"
        FileTextBox.Text = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\" + RPTNAME + ext
    End Sub

    Private Sub QueryButton_Click(sender As Object, e As EventArgs) Handles QueryButton.Click
        CheckForIllegalCrossThreadCalls = False
        vsql = "select *
from master with(nolock)
where master.account=@account"

        QueryThread = New System.Threading.Thread(AddressOf RunQuery
)
        QueryThread.Start()
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
    Public Sub RunQuery()

        Try
            Cursor = Cursors.WaitCursor
            Dim cmd As New SqlCommand(vsql)
            Dim cn As New SqlConnection With {
                .ConnectionString = My.Settings.collect2000ConnectionString
            }
            Dim da1 As SqlDataAdapter = New SqlDataAdapter(vsql, cn.ConnectionString)
            da1.SelectCommand.CommandTimeout = 1800
            da1.SelectCommand.Parameters.AddWithValue("@account", AccountBox.Text)
            dt1.Clear()
            ExportButton.Enabled = False
            cn.Open()
            da1.Fill(dt1)
            da1.Dispose()
            cn.Close()


            ExportButton.Enabled = True


        Catch ex As System.Exception
            Cursor = Cursors.Default
            MessageBox.Show(ex.Message)
            ExportButton.Enabled = False
        End Try
        Cursor = Cursors.Default

    End Sub

    Private Sub ExportButton_Click(sender As Object, e As EventArgs) Handles ExportButton.Click

    End Sub
End Class