<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ReportData
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container()
		Me.RptDataGridView = New System.Windows.Forms.DataGridView()
		Me.ActivityTextBox = New System.Windows.Forms.TextBox()
		Me.ElapseTextBox = New System.Windows.Forms.TextBox()
		Me.ExportButton = New System.Windows.Forms.Button()
		Me.GenerateButton = New System.Windows.Forms.Button()
		Me.FileButton = New System.Windows.Forms.Button()
		Me.StartLabel = New System.Windows.Forms.Label()
		Me.FileTextBox = New System.Windows.Forms.TextBox()
		Me.EndLabel = New System.Windows.Forms.Label()
		Me.enddate = New System.Windows.Forms.DateTimePicker()
		Me.StartDate = New System.Windows.Forms.DateTimePicker()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.GroupBox2 = New System.Windows.Forms.GroupBox()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.StopWatch = New System.Windows.Forms.Timer(Me.components)
		Me.GroupBox3 = New System.Windows.Forms.GroupBox()
		Me.RadioButton2 = New System.Windows.Forms.RadioButton()
		Me.RadioButton1 = New System.Windows.Forms.RadioButton()
		CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox1.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		Me.GroupBox3.SuspendLayout()
		Me.SuspendLayout()
		'
		'RptDataGridView
		'
		Me.RptDataGridView.AllowUserToAddRows = False
		Me.RptDataGridView.AllowUserToDeleteRows = False
		Me.RptDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
		Me.RptDataGridView.BackgroundColor = System.Drawing.SystemColors.Control
		Me.RptDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.RptDataGridView.Dock = System.Windows.Forms.DockStyle.Fill
		Me.RptDataGridView.Location = New System.Drawing.Point(3, 21)
		Me.RptDataGridView.Name = "RptDataGridView"
		Me.RptDataGridView.ReadOnly = True
		Me.RptDataGridView.Size = New System.Drawing.Size(951, 308)
		Me.RptDataGridView.TabIndex = 64
		'
		'ActivityTextBox
		'
		Me.ActivityTextBox.BackColor = System.Drawing.SystemColors.Info
		Me.ActivityTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ActivityTextBox.ForeColor = System.Drawing.SystemColors.InfoText
		Me.ActivityTextBox.Location = New System.Drawing.Point(6, 19)
		Me.ActivityTextBox.Name = "ActivityTextBox"
		Me.ActivityTextBox.ReadOnly = True
		Me.ActivityTextBox.Size = New System.Drawing.Size(190, 23)
		Me.ActivityTextBox.TabIndex = 63
		'
		'ElapseTextBox
		'
		Me.ElapseTextBox.BackColor = System.Drawing.SystemColors.Info
		Me.ElapseTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ElapseTextBox.ForeColor = System.Drawing.SystemColors.InfoText
		Me.ElapseTextBox.Location = New System.Drawing.Point(6, 53)
		Me.ElapseTextBox.Name = "ElapseTextBox"
		Me.ElapseTextBox.ReadOnly = True
		Me.ElapseTextBox.Size = New System.Drawing.Size(190, 23)
		Me.ElapseTextBox.TabIndex = 62
		'
		'ExportButton
		'
		Me.ExportButton.Enabled = False
		Me.ExportButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ExportButton.Location = New System.Drawing.Point(25, 101)
		Me.ExportButton.Name = "ExportButton"
		Me.ExportButton.Size = New System.Drawing.Size(100, 23)
		Me.ExportButton.TabIndex = 61
		Me.ExportButton.Text = "Export Report"
		Me.ExportButton.UseVisualStyleBackColor = True
		'
		'GenerateButton
		'
		Me.GenerateButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GenerateButton.Location = New System.Drawing.Point(25, 61)
		Me.GenerateButton.Name = "GenerateButton"
		Me.GenerateButton.Size = New System.Drawing.Size(100, 23)
		Me.GenerateButton.TabIndex = 60
		Me.GenerateButton.Text = "Generate Report"
		Me.GenerateButton.UseVisualStyleBackColor = True
		'
		'FileButton
		'
		Me.FileButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FileButton.Location = New System.Drawing.Point(25, 21)
		Me.FileButton.Name = "FileButton"
		Me.FileButton.Size = New System.Drawing.Size(100, 23)
		Me.FileButton.TabIndex = 59
		Me.FileButton.Text = "Output File"
		Me.FileButton.UseVisualStyleBackColor = True
		'
		'StartLabel
		'
		Me.StartLabel.AutoSize = True
		Me.StartLabel.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartLabel.Location = New System.Drawing.Point(178, 69)
		Me.StartLabel.Name = "StartLabel"
		Me.StartLabel.Size = New System.Drawing.Size(69, 17)
		Me.StartLabel.TabIndex = 58
		Me.StartLabel.Text = "Start Date"
		'
		'FileTextBox
		'
		Me.FileTextBox.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FileTextBox.Location = New System.Drawing.Point(131, 21)
		Me.FileTextBox.Name = "FileTextBox"
		Me.FileTextBox.Size = New System.Drawing.Size(588, 25)
		Me.FileTextBox.TabIndex = 1
		'
		'EndLabel
		'
		Me.EndLabel.AutoSize = True
		Me.EndLabel.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.EndLabel.Location = New System.Drawing.Point(178, 105)
		Me.EndLabel.Name = "EndLabel"
		Me.EndLabel.Size = New System.Drawing.Size(64, 17)
		Me.EndLabel.TabIndex = 56
		Me.EndLabel.Text = "End Date"
		'
		'enddate
		'
		Me.enddate.CustomFormat = "yyyy-MM-dd 23:59:59.000"
		Me.enddate.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.enddate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.enddate.Location = New System.Drawing.Point(260, 102)
		Me.enddate.Name = "enddate"
		Me.enddate.Size = New System.Drawing.Size(200, 25)
		Me.enddate.TabIndex = 55
		'
		'StartDate
		'
		Me.StartDate.CustomFormat = "yyyy-MM-dd 00:00:00.000"
		Me.StartDate.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.StartDate.Location = New System.Drawing.Point(260, 68)
		Me.StartDate.Name = "StartDate"
		Me.StartDate.Size = New System.Drawing.Size(200, 25)
		Me.StartDate.TabIndex = 54
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.ActivityTextBox)
		Me.GroupBox1.Controls.Add(Me.ElapseTextBox)
		Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox1.Location = New System.Drawing.Point(488, 49)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(206, 87)
		Me.GroupBox1.TabIndex = 65
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Activity"
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.RptDataGridView)
		Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox2.Location = New System.Drawing.Point(23, 134)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(957, 332)
		Me.GroupBox2.TabIndex = 66
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Report Data"
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.CheckFileExists = False
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		Me.OpenFileDialog1.Filter = """CSV files (*.csv)|*.csv|Excel Files (*.xlsx)|*.xlsx"""
		Me.OpenFileDialog1.ValidateNames = False
		'
		'StopWatch
		'
		Me.StopWatch.Interval = 1
		'
		'GroupBox3
		'
		Me.GroupBox3.Controls.Add(Me.RadioButton2)
		Me.GroupBox3.Controls.Add(Me.RadioButton1)
		Me.GroupBox3.Location = New System.Drawing.Point(748, 21)
		Me.GroupBox3.Name = "GroupBox3"
		Me.GroupBox3.Size = New System.Drawing.Size(200, 100)
		Me.GroupBox3.TabIndex = 67
		Me.GroupBox3.TabStop = False
		Me.GroupBox3.Text = "File Type"
		'
		'RadioButton2
		'
		Me.RadioButton2.AutoSize = True
		Me.RadioButton2.Location = New System.Drawing.Point(7, 40)
		Me.RadioButton2.Name = "RadioButton2"
		Me.RadioButton2.Size = New System.Drawing.Size(70, 17)
		Me.RadioButton2.TabIndex = 1
		Me.RadioButton2.TabStop = True
		Me.RadioButton2.Text = "Excel File"
		Me.RadioButton2.UseVisualStyleBackColor = True
		'
		'RadioButton1
		'
		Me.RadioButton1.AutoSize = True
		Me.RadioButton1.Location = New System.Drawing.Point(7, 20)
		Me.RadioButton1.Name = "RadioButton1"
		Me.RadioButton1.Size = New System.Drawing.Size(62, 17)
		Me.RadioButton1.TabIndex = 0
		Me.RadioButton1.TabStop = True
		Me.RadioButton1.Text = "CSV file"
		Me.RadioButton1.UseVisualStyleBackColor = True
		'
		'ReportData
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.Controls.Add(Me.GroupBox3)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.ExportButton)
		Me.Controls.Add(Me.GenerateButton)
		Me.Controls.Add(Me.FileButton)
		Me.Controls.Add(Me.StartLabel)
		Me.Controls.Add(Me.FileTextBox)
		Me.Controls.Add(Me.EndLabel)
		Me.Controls.Add(Me.enddate)
		Me.Controls.Add(Me.StartDate)
		Me.Name = "ReportData"
		Me.Size = New System.Drawing.Size(1564, 495)
		CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox3.ResumeLayout(False)
		Me.GroupBox3.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents RptDataGridView As DataGridView
    Friend WithEvents ActivityTextBox As TextBox
    Friend WithEvents ElapseTextBox As TextBox
    Friend WithEvents ExportButton As Button
    Friend WithEvents GenerateButton As Button
    Friend WithEvents FileButton As Button
    Friend WithEvents StartLabel As Label
    Friend WithEvents FileTextBox As TextBox
    Friend WithEvents EndLabel As Label
    Friend WithEvents enddate As DateTimePicker
    Friend WithEvents StartDate As DateTimePicker
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents StopWatch As Timer
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents RadioButton2 As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
End Class
