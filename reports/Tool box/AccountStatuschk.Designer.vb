<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AccountStatuschk
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
		Me.StartLabel = New System.Windows.Forms.Label()
		Me.EndLabel = New System.Windows.Forms.Label()
		Me.enddate = New System.Windows.Forms.DateTimePicker()
		Me.StartDate = New System.Windows.Forms.DateTimePicker()
		Me.DataGridView1 = New System.Windows.Forms.DataGridView()
		Me.Lookupbtn = New System.Windows.Forms.Button()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.AccountNumber = New System.Windows.Forms.TextBox()
		Me.StatCalc = New System.Windows.Forms.TextBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.allnotesbtn = New System.Windows.Forms.RadioButton()
		Me.statusnotesbtn = New System.Windows.Forms.RadioButton()
		Me.Button1 = New System.Windows.Forms.Button()
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'StartLabel
		'
		Me.StartLabel.AutoSize = True
		Me.StartLabel.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartLabel.Location = New System.Drawing.Point(280, 22)
		Me.StartLabel.Name = "StartLabel"
		Me.StartLabel.Size = New System.Drawing.Size(69, 17)
		Me.StartLabel.TabIndex = 62
		Me.StartLabel.Text = "Start Date"
		'
		'EndLabel
		'
		Me.EndLabel.AutoSize = True
		Me.EndLabel.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.EndLabel.Location = New System.Drawing.Point(280, 58)
		Me.EndLabel.Name = "EndLabel"
		Me.EndLabel.Size = New System.Drawing.Size(64, 17)
		Me.EndLabel.TabIndex = 61
		Me.EndLabel.Text = "End Date"
		'
		'enddate
		'
		Me.enddate.CustomFormat = "yyyy-MM-dd 23:59:59.000"
		Me.enddate.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.enddate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.enddate.Location = New System.Drawing.Point(362, 55)
		Me.enddate.Name = "enddate"
		Me.enddate.Size = New System.Drawing.Size(200, 25)
		Me.enddate.TabIndex = 60
		'
		'StartDate
		'
		Me.StartDate.CustomFormat = "yyyy-MM-dd 00:00:00.000"
		Me.StartDate.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.StartDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.StartDate.Location = New System.Drawing.Point(362, 21)
		Me.StartDate.Name = "StartDate"
		Me.StartDate.Size = New System.Drawing.Size(200, 25)
		Me.StartDate.TabIndex = 59
		'
		'DataGridView1
		'
		Me.DataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
		Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.DisplayedCells
		Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DataGridView1.Location = New System.Drawing.Point(46, 165)
		Me.DataGridView1.MultiSelect = False
		Me.DataGridView1.Name = "DataGridView1"
		Me.DataGridView1.Size = New System.Drawing.Size(676, 248)
		Me.DataGridView1.TabIndex = 68
		'
		'Lookupbtn
		'
		Me.Lookupbtn.Location = New System.Drawing.Point(46, 58)
		Me.Lookupbtn.Name = "Lookupbtn"
		Me.Lookupbtn.Size = New System.Drawing.Size(75, 23)
		Me.Lookupbtn.TabIndex = 65
		Me.Lookupbtn.Text = "Lookup"
		Me.Lookupbtn.UseVisualStyleBackColor = True
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(43, 28)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(47, 13)
		Me.Label1.TabIndex = 64
		Me.Label1.Text = "Account"
		'
		'AccountNumber
		'
		Me.AccountNumber.Location = New System.Drawing.Point(94, 22)
		Me.AccountNumber.Name = "AccountNumber"
		Me.AccountNumber.Size = New System.Drawing.Size(137, 20)
		Me.AccountNumber.TabIndex = 63
		'
		'StatCalc
		'
		Me.StatCalc.Location = New System.Drawing.Point(145, 139)
		Me.StatCalc.Name = "StatCalc"
		Me.StatCalc.ReadOnly = True
		Me.StatCalc.Size = New System.Drawing.Size(100, 20)
		Me.StatCalc.TabIndex = 69
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(49, 146)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(90, 13)
		Me.Label2.TabIndex = 70
		Me.Label2.Text = "Calculated Status"
		'
		'allnotesbtn
		'
		Me.allnotesbtn.AutoSize = True
		Me.allnotesbtn.Location = New System.Drawing.Point(37, 88)
		Me.allnotesbtn.Name = "allnotesbtn"
		Me.allnotesbtn.Size = New System.Drawing.Size(67, 17)
		Me.allnotesbtn.TabIndex = 71
		Me.allnotesbtn.TabStop = True
		Me.allnotesbtn.Text = "All Notes"
		Me.allnotesbtn.UseVisualStyleBackColor = True
		'
		'statusnotesbtn
		'
		Me.statusnotesbtn.AutoSize = True
		Me.statusnotesbtn.Location = New System.Drawing.Point(123, 88)
		Me.statusnotesbtn.Name = "statusnotesbtn"
		Me.statusnotesbtn.Size = New System.Drawing.Size(86, 17)
		Me.statusnotesbtn.TabIndex = 72
		Me.statusnotesbtn.TabStop = True
		Me.statusnotesbtn.Text = "Status Notes"
		Me.statusnotesbtn.UseVisualStyleBackColor = True
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(300, 139)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(130, 23)
		Me.Button1.TabIndex = 73
		Me.Button1.Text = "Set This Status"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'AccountStatuschk
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.Button1)
		Me.Controls.Add(Me.statusnotesbtn)
		Me.Controls.Add(Me.allnotesbtn)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.StatCalc)
		Me.Controls.Add(Me.DataGridView1)
		Me.Controls.Add(Me.Lookupbtn)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.AccountNumber)
		Me.Controls.Add(Me.StartLabel)
		Me.Controls.Add(Me.EndLabel)
		Me.Controls.Add(Me.enddate)
		Me.Controls.Add(Me.StartDate)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "AccountStatuschk"
		Me.Text = "Account Status Check"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents StartLabel As Label
	Friend WithEvents EndLabel As Label
	Friend WithEvents enddate As DateTimePicker
	Friend WithEvents StartDate As DateTimePicker
	Friend WithEvents DataGridView1 As DataGridView
	Friend WithEvents Lookupbtn As Button
	Friend WithEvents Label1 As Label
	Friend WithEvents AccountNumber As TextBox
	Friend WithEvents StatCalc As TextBox
	Friend WithEvents Label2 As Label
	Friend WithEvents allnotesbtn As RadioButton
	Friend WithEvents statusnotesbtn As RadioButton
	Friend WithEvents Button1 As Button
End Class
