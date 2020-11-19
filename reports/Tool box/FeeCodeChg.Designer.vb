<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FeeCodeChg
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
		Me.DataGridView1 = New System.Windows.Forms.DataGridView()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.GroupBox2 = New System.Windows.Forms.GroupBox()
		Me.prod = New System.Windows.Forms.RadioButton()
		Me.testdb = New System.Windows.Forms.RadioButton()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.ActivityTextBox = New System.Windows.Forms.TextBox()
		Me.ElapseTextBox = New System.Windows.Forms.TextBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.ACount = New System.Windows.Forms.TextBox()
		Me.StopWatch = New System.Windows.Forms.Timer(Me.components)
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.ComboBox1 = New System.Windows.Forms.ComboBox()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.ComboBox2 = New System.Windows.Forms.ComboBox()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.MakeChange = New System.Windows.Forms.Button()
		Me.Initials = New System.Windows.Forms.TextBox()
		Me.Label2 = New System.Windows.Forms.Label()
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox2.SuspendLayout()
		Me.GroupBox1.SuspendLayout()
		Me.SuspendLayout()
		'
		'DataGridView1
		'
		Me.DataGridView1.AllowUserToAddRows = False
		Me.DataGridView1.AllowUserToDeleteRows = False
		Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
		Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
		Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DataGridView1.Location = New System.Drawing.Point(21, 52)
		Me.DataGridView1.Name = "DataGridView1"
		Me.DataGridView1.ReadOnly = True
		Me.DataGridView1.Size = New System.Drawing.Size(738, 150)
		Me.DataGridView1.TabIndex = 3
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(22, 24)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(75, 23)
		Me.Button1.TabIndex = 2
		Me.Button1.Text = "Load Excel"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.prod)
		Me.GroupBox2.Controls.Add(Me.testdb)
		Me.GroupBox2.Location = New System.Drawing.Point(603, 211)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(156, 78)
		Me.GroupBox2.TabIndex = 82
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Database"
		'
		'prod
		'
		Me.prod.AutoSize = True
		Me.prod.Location = New System.Drawing.Point(6, 26)
		Me.prod.Name = "prod"
		Me.prod.Size = New System.Drawing.Size(131, 17)
		Me.prod.TabIndex = 76
		Me.prod.TabStop = True
		Me.prod.Text = "fbpa1sql2 (Production)"
		Me.prod.UseVisualStyleBackColor = True
		'
		'testdb
		'
		Me.testdb.AutoSize = True
		Me.testdb.Location = New System.Drawing.Point(6, 49)
		Me.testdb.Name = "testdb"
		Me.testdb.Size = New System.Drawing.Size(114, 17)
		Me.testdb.TabIndex = 77
		Me.testdb.TabStop = True
		Me.testdb.Text = "fbpa1sglapril (Test)"
		Me.testdb.UseVisualStyleBackColor = True
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.ActivityTextBox)
		Me.GroupBox1.Controls.Add(Me.ElapseTextBox)
		Me.GroupBox1.Font = New System.Drawing.Font("Times New Roman", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox1.Location = New System.Drawing.Point(425, 292)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(334, 87)
		Me.GroupBox1.TabIndex = 81
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
		Me.ActivityTextBox.Size = New System.Drawing.Size(316, 23)
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
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(138, 210)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(52, 13)
		Me.Label1.TabIndex = 80
		Me.Label1.Text = "Accounts"
		'
		'ACount
		'
		Me.ACount.Location = New System.Drawing.Point(208, 207)
		Me.ACount.Name = "ACount"
		Me.ACount.ReadOnly = True
		Me.ACount.Size = New System.Drawing.Size(147, 20)
		Me.ACount.TabIndex = 79
		'
		'StopWatch
		'
		Me.StopWatch.Interval = 1
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		'
		'ComboBox1
		'
		Me.ComboBox1.FormattingEnabled = True
		Me.ComboBox1.Location = New System.Drawing.Point(208, 233)
		Me.ComboBox1.Name = "ComboBox1"
		Me.ComboBox1.Size = New System.Drawing.Size(147, 21)
		Me.ComboBox1.TabIndex = 84
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(119, 239)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(71, 13)
		Me.Label3.TabIndex = 83
		Me.Label3.Text = "File # Column"
		'
		'ComboBox2
		'
		Me.ComboBox2.FormattingEnabled = True
		Me.ComboBox2.Location = New System.Drawing.Point(208, 260)
		Me.ComboBox2.Name = "ComboBox2"
		Me.ComboBox2.Size = New System.Drawing.Size(147, 21)
		Me.ComboBox2.TabIndex = 86
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Location = New System.Drawing.Point(99, 266)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(91, 13)
		Me.Label4.TabIndex = 85
		Me.Label4.Text = "Fee Code Column"
		'
		'MakeChange
		'
		Me.MakeChange.AutoEllipsis = True
		Me.MakeChange.Enabled = False
		Me.MakeChange.Location = New System.Drawing.Point(208, 356)
		Me.MakeChange.Name = "MakeChange"
		Me.MakeChange.Size = New System.Drawing.Size(108, 23)
		Me.MakeChange.TabIndex = 87
		Me.MakeChange.Text = "Make Change"
		Me.MakeChange.UseVisualStyleBackColor = True
		'
		'Initials
		'
		Me.Initials.Location = New System.Drawing.Point(208, 287)
		Me.Initials.Name = "Initials"
		Me.Initials.Size = New System.Drawing.Size(100, 20)
		Me.Initials.TabIndex = 88
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(130, 294)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(60, 13)
		Me.Label2.TabIndex = 89
		Me.Label2.Text = "User initials"
		'
		'FeeCodeChg
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Initials)
		Me.Controls.Add(Me.MakeChange)
		Me.Controls.Add(Me.ComboBox2)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.ComboBox1)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.ACount)
		Me.Controls.Add(Me.DataGridView1)
		Me.Controls.Add(Me.Button1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "FeeCodeChg"
		Me.Text = "Fee Code Maintenance"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox2.PerformLayout()
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents DataGridView1 As DataGridView
	Friend WithEvents Button1 As Button
	Friend WithEvents GroupBox2 As GroupBox
	Friend WithEvents prod As RadioButton
	Friend WithEvents testdb As RadioButton
	Friend WithEvents GroupBox1 As GroupBox
	Friend WithEvents ActivityTextBox As TextBox
	Friend WithEvents ElapseTextBox As TextBox
	Friend WithEvents Label1 As Label
	Friend WithEvents ACount As TextBox
	Friend WithEvents StopWatch As Timer
	Friend WithEvents OpenFileDialog1 As OpenFileDialog
	Friend WithEvents ComboBox1 As ComboBox
	Friend WithEvents Label3 As Label
	Friend WithEvents ComboBox2 As ComboBox
	Friend WithEvents Label4 As Label
	Friend WithEvents MakeChange As Button
	Friend WithEvents Initials As TextBox
	Friend WithEvents Label2 As Label
End Class
