<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ExeterSifPif
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.StopWatch = New System.Windows.Forms.Timer(Me.components)
		Me.ReportData1 = New reports.ReportData()
		Me.SuspendLayout()
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.CheckFileExists = False
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		Me.OpenFileDialog1.Filter = "*.xlsx|*.csv"
		Me.OpenFileDialog1.ValidateNames = False
		'
		'StopWatch
		'
		Me.StopWatch.Interval = 1
		'
		'ReportData1
		'
		Me.ReportData1.Location = New System.Drawing.Point(13, 13)
		Me.ReportData1.Name = "ReportData1"
		Me.ReportData1.Size = New System.Drawing.Size(1564, 495)
		Me.ReportData1.TabIndex = 0
		'
		'ExeterSifPif
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.CausesValidation = False
		Me.ClientSize = New System.Drawing.Size(1372, 450)
		Me.ControlBox = False
		Me.Controls.Add(Me.ReportData1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "ExeterSifPif"
		Me.ShowIcon = False
		Me.Text = "ExeterSifPif"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.ResumeLayout(False)

	End Sub
	Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents StopWatch As Timer
    Friend WithEvents ReportData1 As ReportData
End Class
