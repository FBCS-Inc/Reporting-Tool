<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DateMaintenance
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
		Me.filenumber = New System.Windows.Forms.TextBox()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.ClosedDate = New System.Windows.Forms.DateTimePicker()
		Me.ReturnedDate = New System.Windows.Forms.DateTimePicker()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.closedbox = New System.Windows.Forms.TextBox()
		Me.returnedbox = New System.Windows.Forms.TextBox()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label5 = New System.Windows.Forms.Label()
		Me.Button2 = New System.Windows.Forms.Button()
		Me.Button3 = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'filenumber
		'
		Me.filenumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.filenumber.Location = New System.Drawing.Point(179, 24)
		Me.filenumber.Name = "filenumber"
		Me.filenumber.Size = New System.Drawing.Size(129, 26)
		Me.filenumber.TabIndex = 0
		'
		'Button1
		'
		Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Button1.Location = New System.Drawing.Point(331, 22)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(178, 28)
		Me.Button1.TabIndex = 1
		Me.Button1.Text = "Look Up File Number"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'ClosedDate
		'
		Me.ClosedDate.CustomFormat = "yyyy-MM-dd"
		Me.ClosedDate.Enabled = False
		Me.ClosedDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ClosedDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.ClosedDate.Location = New System.Drawing.Point(447, 80)
		Me.ClosedDate.Name = "ClosedDate"
		Me.ClosedDate.Size = New System.Drawing.Size(200, 26)
		Me.ClosedDate.TabIndex = 2
		'
		'ReturnedDate
		'
		Me.ReturnedDate.CustomFormat = "yyyy-MM-dd"
		Me.ReturnedDate.Enabled = False
		Me.ReturnedDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ReturnedDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
		Me.ReturnedDate.Location = New System.Drawing.Point(447, 124)
		Me.ReturnedDate.Name = "ReturnedDate"
		Me.ReturnedDate.Size = New System.Drawing.Size(200, 26)
		Me.ReturnedDate.TabIndex = 3
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Enabled = False
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(76, 83)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(97, 20)
		Me.Label1.TabIndex = 4
		Me.Label1.Text = "Closed Date"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Enabled = False
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(58, 124)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(115, 20)
		Me.Label2.TabIndex = 5
		Me.Label2.Text = "Returned Date"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(79, 30)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(94, 20)
		Me.Label3.TabIndex = 6
		Me.Label3.Text = "File Number"
		'
		'closedbox
		'
		Me.closedbox.Enabled = False
		Me.closedbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.closedbox.Location = New System.Drawing.Point(179, 80)
		Me.closedbox.Name = "closedbox"
		Me.closedbox.Size = New System.Drawing.Size(129, 26)
		Me.closedbox.TabIndex = 7
		'
		'returnedbox
		'
		Me.returnedbox.Enabled = False
		Me.returnedbox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.returnedbox.Location = New System.Drawing.Point(179, 124)
		Me.returnedbox.Name = "returnedbox"
		Me.returnedbox.Size = New System.Drawing.Size(129, 26)
		Me.returnedbox.TabIndex = 8
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Enabled = False
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(325, 86)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(91, 20)
		Me.Label4.TabIndex = 9
		Me.Label4.Text = "Change To:"
		'
		'Label5
		'
		Me.Label5.AutoSize = True
		Me.Label5.Enabled = False
		Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Location = New System.Drawing.Point(327, 130)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(91, 20)
		Me.Label5.TabIndex = 10
		Me.Label5.Text = "Change To:"
		'
		'Button2
		'
		Me.Button2.Enabled = False
		Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Button2.Location = New System.Drawing.Point(673, 80)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(168, 26)
		Me.Button2.TabIndex = 11
		Me.Button2.Text = "Save Closed Date"
		Me.Button2.UseVisualStyleBackColor = True
		'
		'Button3
		'
		Me.Button3.Enabled = False
		Me.Button3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Button3.Location = New System.Drawing.Point(673, 124)
		Me.Button3.Name = "Button3"
		Me.Button3.Size = New System.Drawing.Size(168, 26)
		Me.Button3.TabIndex = 12
		Me.Button3.Text = "Save Returned Date"
		Me.Button3.UseVisualStyleBackColor = True
		'
		'DateMaintenance
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1076, 450)
		Me.ControlBox = False
		Me.Controls.Add(Me.Button3)
		Me.Controls.Add(Me.Button2)
		Me.Controls.Add(Me.Label5)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.returnedbox)
		Me.Controls.Add(Me.closedbox)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.ReturnedDate)
		Me.Controls.Add(Me.ClosedDate)
		Me.Controls.Add(Me.Button1)
		Me.Controls.Add(Me.filenumber)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "DateMaintenance"
		Me.Text = "Date Maintenance"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents filenumber As TextBox
	Friend WithEvents Button1 As Button
	Friend WithEvents ClosedDate As DateTimePicker
	Friend WithEvents ReturnedDate As DateTimePicker
	Friend WithEvents Label1 As Label
	Friend WithEvents Label2 As Label
	Friend WithEvents Label3 As Label
	Friend WithEvents closedbox As TextBox
	Friend WithEvents returnedbox As TextBox
	Friend WithEvents Label4 As Label
	Friend WithEvents Label5 As Label
	Friend WithEvents Button2 As Button
	Friend WithEvents Button3 As Button
End Class
