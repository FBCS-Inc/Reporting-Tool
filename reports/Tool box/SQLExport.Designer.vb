<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SQLExport
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.StopWatch = New System.Windows.Forms.Timer(Me.components)
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.ActivityTextBox = New System.Windows.Forms.TextBox()
		Me.ElapseTextBox = New System.Windows.Forms.TextBox()
		Me.GroupBox3 = New System.Windows.Forms.GroupBox()
		Me.excelfiletype = New System.Windows.Forms.RadioButton()
		Me.RadioButton1 = New System.Windows.Forms.RadioButton()
		Me.FileButton = New System.Windows.Forms.Button()
		Me.FileTextBox = New System.Windows.Forms.TextBox()
		Me.SQLBox = New System.Windows.Forms.RichTextBox()
		Me.QueryButton = New System.Windows.Forms.Button()
		Me.ExportButton = New System.Windows.Forms.Button()
		Me.GroupBox2 = New System.Windows.Forms.GroupBox()
		Me.RptDataGridView = New System.Windows.Forms.DataGridView()
		Me.GroupBox4 = New System.Windows.Forms.GroupBox()
		Me.label1 = New System.Windows.Forms.Label()
		Me.Fmtbox = New System.Windows.Forms.TextBox()
		Me.LimitBox = New System.Windows.Forms.TextBox()
		Me.SheetLimit = New System.Windows.Forms.RadioButton()
		Me.sheetall = New System.Windows.Forms.RadioButton()
		Me.CustList1 = New reports.CustList()
		Me.GroupBox1.SuspendLayout()
		Me.GroupBox3.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox4.SuspendLayout()
		Me.SuspendLayout()
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
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.ActivityTextBox)
		Me.GroupBox1.Controls.Add(Me.ElapseTextBox)
		Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox1.Location = New System.Drawing.Point(727, 89)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(206, 87)
		Me.GroupBox1.TabIndex = 66
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Activity"
		'
		'ActivityTextBox
		'
		Me.ActivityTextBox.BackColor = System.Drawing.SystemColors.Info
		Me.ActivityTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ActivityTextBox.ForeColor = System.Drawing.SystemColors.InfoText
		Me.ActivityTextBox.Location = New System.Drawing.Point(6, 21)
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
		'GroupBox3
		'
		Me.GroupBox3.Controls.Add(Me.excelfiletype)
		Me.GroupBox3.Controls.Add(Me.RadioButton1)
		Me.GroupBox3.Location = New System.Drawing.Point(727, 12)
		Me.GroupBox3.Name = "GroupBox3"
		Me.GroupBox3.Size = New System.Drawing.Size(200, 72)
		Me.GroupBox3.TabIndex = 68
		Me.GroupBox3.TabStop = False
		Me.GroupBox3.Text = "File Type"
		'
		'excelfiletype
		'
		Me.excelfiletype.AutoSize = True
		Me.excelfiletype.Location = New System.Drawing.Point(7, 40)
		Me.excelfiletype.Name = "excelfiletype"
		Me.excelfiletype.Size = New System.Drawing.Size(70, 17)
		Me.excelfiletype.TabIndex = 1
		Me.excelfiletype.TabStop = True
		Me.excelfiletype.Text = "Excel File"
		Me.excelfiletype.UseVisualStyleBackColor = True
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
		'FileButton
		'
		Me.FileButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FileButton.Location = New System.Drawing.Point(12, 12)
		Me.FileButton.Name = "FileButton"
		Me.FileButton.Size = New System.Drawing.Size(100, 23)
		Me.FileButton.TabIndex = 70
		Me.FileButton.Text = "Output File"
		Me.FileButton.UseVisualStyleBackColor = True
		'
		'FileTextBox
		'
		Me.FileTextBox.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FileTextBox.Location = New System.Drawing.Point(118, 12)
		Me.FileTextBox.Name = "FileTextBox"
		Me.FileTextBox.Size = New System.Drawing.Size(588, 25)
		Me.FileTextBox.TabIndex = 69
		'
		'SQLBox
		'
		Me.SQLBox.Location = New System.Drawing.Point(118, 104)
		Me.SQLBox.Name = "SQLBox"
		Me.SQLBox.Size = New System.Drawing.Size(588, 154)
		Me.SQLBox.TabIndex = 71
		Me.SQLBox.Text = ""
		'
		'QueryButton
		'
		Me.QueryButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.QueryButton.Location = New System.Drawing.Point(12, 105)
		Me.QueryButton.Name = "QueryButton"
		Me.QueryButton.Size = New System.Drawing.Size(100, 23)
		Me.QueryButton.TabIndex = 72
		Me.QueryButton.Text = "Run Query"
		Me.QueryButton.UseVisualStyleBackColor = True
		'
		'ExportButton
		'
		Me.ExportButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ExportButton.Location = New System.Drawing.Point(15, 142)
		Me.ExportButton.Name = "ExportButton"
		Me.ExportButton.Size = New System.Drawing.Size(100, 23)
		Me.ExportButton.TabIndex = 73
		Me.ExportButton.Text = "Export"
		Me.ExportButton.UseVisualStyleBackColor = True
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.RptDataGridView)
		Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox2.Location = New System.Drawing.Point(12, 336)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(957, 332)
		Me.GroupBox2.TabIndex = 74
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Report Data"
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
		'GroupBox4
		'
		Me.GroupBox4.Controls.Add(Me.label1)
		Me.GroupBox4.Controls.Add(Me.Fmtbox)
		Me.GroupBox4.Controls.Add(Me.LimitBox)
		Me.GroupBox4.Controls.Add(Me.SheetLimit)
		Me.GroupBox4.Controls.Add(Me.sheetall)
		Me.GroupBox4.Location = New System.Drawing.Point(118, 264)
		Me.GroupBox4.Name = "GroupBox4"
		Me.GroupBox4.Size = New System.Drawing.Size(483, 66)
		Me.GroupBox4.TabIndex = 75
		Me.GroupBox4.TabStop = False
		Me.GroupBox4.Text = "Sheet Properties"
		Me.GroupBox4.Visible = False
		'
		'label1
		'
		Me.label1.AutoSize = True
		Me.label1.Location = New System.Drawing.Point(236, 20)
		Me.label1.Name = "label1"
		Me.label1.Size = New System.Drawing.Size(42, 13)
		Me.label1.TabIndex = 4
		Me.label1.Text = "Format "
		'
		'Fmtbox
		'
		Me.Fmtbox.Location = New System.Drawing.Point(284, 17)
		Me.Fmtbox.Name = "Fmtbox"
		Me.Fmtbox.Size = New System.Drawing.Size(193, 20)
		Me.Fmtbox.TabIndex = 3
		'
		'LimitBox
		'
		Me.LimitBox.Location = New System.Drawing.Point(86, 40)
		Me.LimitBox.Name = "LimitBox"
		Me.LimitBox.Size = New System.Drawing.Size(100, 20)
		Me.LimitBox.TabIndex = 2
		'
		'SheetLimit
		'
		Me.SheetLimit.AutoSize = True
		Me.SheetLimit.Location = New System.Drawing.Point(7, 43)
		Me.SheetLimit.Name = "SheetLimit"
		Me.SheetLimit.Size = New System.Drawing.Size(58, 17)
		Me.SheetLimit.TabIndex = 1
		Me.SheetLimit.TabStop = True
		Me.SheetLimit.Text = "Limit to"
		Me.SheetLimit.UseVisualStyleBackColor = True
		'
		'sheetall
		'
		Me.sheetall.AutoSize = True
		Me.sheetall.Location = New System.Drawing.Point(7, 20)
		Me.sheetall.Name = "sheetall"
		Me.sheetall.Size = New System.Drawing.Size(70, 17)
		Me.sheetall.TabIndex = 0
		Me.sheetall.TabStop = True
		Me.sheetall.Text = "ANo Limit"
		Me.sheetall.UseVisualStyleBackColor = True
		'
		'CustList1
		'
		Me.CustList1.Location = New System.Drawing.Point(15, 43)
		Me.CustList1.Name = "CustList1"
		Me.CustList1.Size = New System.Drawing.Size(602, 55)
		Me.CustList1.TabIndex = 79
		'
		'SQLExport
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1269, 866)
		Me.ControlBox = False
		Me.Controls.Add(Me.CustList1)
		Me.Controls.Add(Me.GroupBox4)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.ExportButton)
		Me.Controls.Add(Me.QueryButton)
		Me.Controls.Add(Me.SQLBox)
		Me.Controls.Add(Me.FileButton)
		Me.Controls.Add(Me.FileTextBox)
		Me.Controls.Add(Me.GroupBox3)
		Me.Controls.Add(Me.GroupBox1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "SQLExport"
		Me.Text = "SQLExport"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.GroupBox3.ResumeLayout(False)
		Me.GroupBox3.PerformLayout()
		Me.GroupBox2.ResumeLayout(False)
		CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox4.ResumeLayout(False)
		Me.GroupBox4.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents StopWatch As Timer
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents ActivityTextBox As TextBox
    Friend WithEvents ElapseTextBox As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents excelfiletype As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents FileButton As Button
    Friend WithEvents FileTextBox As TextBox
    Friend WithEvents SQLBox As RichTextBox
    Friend WithEvents QueryButton As Button
    Friend WithEvents ExportButton As Button
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents RptDataGridView As DataGridView
    Friend WithEvents GroupBox4 As GroupBox
    Friend WithEvents LimitBox As TextBox
    Friend WithEvents SheetLimit As RadioButton
    Friend WithEvents sheetall As RadioButton
    Friend WithEvents label1 As Label
    Friend WithEvents Fmtbox As TextBox
	Friend WithEvents CustList1 As CustList
End Class
