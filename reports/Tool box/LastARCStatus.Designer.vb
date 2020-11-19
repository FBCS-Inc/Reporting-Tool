<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LastARCStatus
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
		Me.Accountbox = New System.Windows.Forms.TextBox()
		Me.CurrentARC = New System.Windows.Forms.TextBox()
		Me.NewARC = New System.Windows.Forms.TextBox()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label3 = New System.Windows.Forms.Label()
		Me.BeforeDataGridView = New System.Windows.Forms.DataGridView()
		Me.AfterDataGridView = New System.Windows.Forms.DataGridView()
		Me.BeforeButton = New System.Windows.Forms.Button()
		Me.AfterButton = New System.Windows.Forms.Button()
		Me.Filebox = New System.Windows.Forms.TextBox()
		Me.DateBox = New System.Windows.Forms.TextBox()
		Me.Label4 = New System.Windows.Forms.Label()
		Me.Label5 = New System.Windows.Forms.Label()
		Me.prebox = New System.Windows.Forms.TextBox()
		Me.Label6 = New System.Windows.Forms.Label()
		CType(Me.BeforeDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.AfterDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'Accountbox
		'
		Me.Accountbox.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Accountbox.Location = New System.Drawing.Point(393, 38)
		Me.Accountbox.Name = "Accountbox"
		Me.Accountbox.Size = New System.Drawing.Size(167, 26)
		Me.Accountbox.TabIndex = 0
		'
		'CurrentARC
		'
		Me.CurrentARC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.CurrentARC.Location = New System.Drawing.Point(393, 79)
		Me.CurrentARC.Name = "CurrentARC"
		Me.CurrentARC.Size = New System.Drawing.Size(167, 26)
		Me.CurrentARC.TabIndex = 1
		'
		'NewARC
		'
		Me.NewARC.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.NewARC.Location = New System.Drawing.Point(393, 120)
		Me.NewARC.Name = "NewARC"
		Me.NewARC.Size = New System.Drawing.Size(167, 26)
		Me.NewARC.TabIndex = 2
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(300, 45)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(72, 19)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "Account #"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(241, 86)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(131, 19)
		Me.Label2.TabIndex = 4
		Me.Label2.Text = "Current ARC Status"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(257, 127)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(115, 19)
		Me.Label3.TabIndex = 5
		Me.Label3.Text = "New ARC Status"
		'
		'BeforeDataGridView
		'
		Me.BeforeDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.BeforeDataGridView.Location = New System.Drawing.Point(78, 175)
		Me.BeforeDataGridView.Name = "BeforeDataGridView"
		Me.BeforeDataGridView.Size = New System.Drawing.Size(814, 92)
		Me.BeforeDataGridView.TabIndex = 6
		'
		'AfterDataGridView
		'
		Me.AfterDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.AfterDataGridView.Location = New System.Drawing.Point(78, 288)
		Me.AfterDataGridView.Name = "AfterDataGridView"
		Me.AfterDataGridView.Size = New System.Drawing.Size(814, 92)
		Me.AfterDataGridView.TabIndex = 7
		'
		'BeforeButton
		'
		Me.BeforeButton.Enabled = False
		Me.BeforeButton.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.BeforeButton.Location = New System.Drawing.Point(78, 75)
		Me.BeforeButton.Name = "BeforeButton"
		Me.BeforeButton.Size = New System.Drawing.Size(129, 30)
		Me.BeforeButton.TabIndex = 8
		Me.BeforeButton.Text = "Before Condition"
		Me.BeforeButton.UseVisualStyleBackColor = True
		'
		'AfterButton
		'
		Me.AfterButton.Enabled = False
		Me.AfterButton.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AfterButton.Location = New System.Drawing.Point(78, 120)
		Me.AfterButton.Name = "AfterButton"
		Me.AfterButton.Size = New System.Drawing.Size(129, 30)
		Me.AfterButton.TabIndex = 9
		Me.AfterButton.Text = "Make Update"
		Me.AfterButton.UseVisualStyleBackColor = True
		'
		'Filebox
		'
		Me.Filebox.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Filebox.Location = New System.Drawing.Point(725, 38)
		Me.Filebox.Name = "Filebox"
		Me.Filebox.Size = New System.Drawing.Size(167, 26)
		Me.Filebox.TabIndex = 10
		'
		'DateBox
		'
		Me.DateBox.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.DateBox.Location = New System.Drawing.Point(725, 79)
		Me.DateBox.Name = "DateBox"
		Me.DateBox.Size = New System.Drawing.Size(167, 26)
		Me.DateBox.TabIndex = 11
		'
		'Label4
		'
		Me.Label4.AutoSize = True
		Me.Label4.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(676, 41)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(43, 19)
		Me.Label4.TabIndex = 12
		Me.Label4.Text = "File #"
		'
		'Label5
		'
		Me.Label5.AutoSize = True
		Me.Label5.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Location = New System.Drawing.Point(640, 81)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(79, 19)
		Me.Label5.TabIndex = 13
		Me.Label5.Text = "Status Date"
		'
		'prebox
		'
		Me.prebox.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.prebox.Location = New System.Drawing.Point(725, 120)
		Me.prebox.Name = "prebox"
		Me.prebox.Size = New System.Drawing.Size(167, 26)
		Me.prebox.TabIndex = 14
		'
		'Label6
		'
		Me.Label6.AutoSize = True
		Me.Label6.Font = New System.Drawing.Font("Times New Roman", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Location = New System.Drawing.Point(613, 123)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(106, 19)
		Me.Label6.TabIndex = 15
		Me.Label6.Text = "Pre Issue Status"
		'
		'LastARCStatus
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1008, 450)
		Me.ControlBox = False
		Me.Controls.Add(Me.Label6)
		Me.Controls.Add(Me.prebox)
		Me.Controls.Add(Me.Label5)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.DateBox)
		Me.Controls.Add(Me.Filebox)
		Me.Controls.Add(Me.AfterButton)
		Me.Controls.Add(Me.BeforeButton)
		Me.Controls.Add(Me.AfterDataGridView)
		Me.Controls.Add(Me.BeforeDataGridView)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.NewARC)
		Me.Controls.Add(Me.CurrentARC)
		Me.Controls.Add(Me.Accountbox)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "LastARCStatus"
		Me.Text = "LastARCStatus"
		CType(Me.BeforeDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.AfterDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents Accountbox As TextBox
    Friend WithEvents CurrentARC As TextBox
    Friend WithEvents NewARC As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents BeforeDataGridView As DataGridView
    Friend WithEvents AfterDataGridView As DataGridView
    Friend WithEvents BeforeButton As Button
    Friend WithEvents AfterButton As Button
    Friend WithEvents Filebox As TextBox
    Friend WithEvents DateBox As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents prebox As TextBox
    Friend WithEvents Label6 As Label
End Class
