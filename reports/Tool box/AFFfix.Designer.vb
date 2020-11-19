<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class AFFfix
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
		Me.jsonfile = New System.Windows.Forms.TextBox()
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
		Me.jsoninput = New System.Windows.Forms.Button()
		Me.DataGridView1 = New System.Windows.Forms.DataGridView()
		Me.Clean = New System.Windows.Forms.Button()
		Me.rawcount = New System.Windows.Forms.TextBox()
		Me.GroupBox1 = New System.Windows.Forms.GroupBox()
		Me.Label2 = New System.Windows.Forms.Label()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.cleancount = New System.Windows.Forms.TextBox()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.DataGridView2 = New System.Windows.Forms.DataGridView()
		Me.GroupBox2 = New System.Windows.Forms.GroupBox()
		Me.GroupBox3 = New System.Windows.Forms.GroupBox()
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox1.SuspendLayout()
		CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.GroupBox2.SuspendLayout()
		Me.GroupBox3.SuspendLayout()
		Me.SuspendLayout()
		'
		'jsonfile
		'
		Me.jsonfile.Location = New System.Drawing.Point(159, 12)
		Me.jsonfile.Name = "jsonfile"
		Me.jsonfile.Size = New System.Drawing.Size(641, 20)
		Me.jsonfile.TabIndex = 0
		'
		'OpenFileDialog1
		'
		Me.OpenFileDialog1.FileName = "OpenFileDialog1"
		'
		'jsoninput
		'
		Me.jsoninput.Location = New System.Drawing.Point(72, 12)
		Me.jsoninput.Name = "jsoninput"
		Me.jsoninput.Size = New System.Drawing.Size(75, 23)
		Me.jsoninput.TabIndex = 1
		Me.jsoninput.Text = "Select File"
		Me.jsoninput.UseVisualStyleBackColor = True
		'
		'DataGridView1
		'
		Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DataGridView1.Location = New System.Drawing.Point(20, 19)
		Me.DataGridView1.Name = "DataGridView1"
		Me.DataGridView1.ReadOnly = True
		Me.DataGridView1.Size = New System.Drawing.Size(980, 244)
		Me.DataGridView1.TabIndex = 3
		'
		'Clean
		'
		Me.Clean.Location = New System.Drawing.Point(72, 42)
		Me.Clean.Name = "Clean"
		Me.Clean.Size = New System.Drawing.Size(75, 23)
		Me.Clean.TabIndex = 4
		Me.Clean.Text = "Clean File"
		Me.Clean.UseVisualStyleBackColor = True
		'
		'rawcount
		'
		Me.rawcount.Location = New System.Drawing.Point(94, 19)
		Me.rawcount.Name = "rawcount"
		Me.rawcount.ReadOnly = True
		Me.rawcount.Size = New System.Drawing.Size(100, 20)
		Me.rawcount.TabIndex = 5
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.Label2)
		Me.GroupBox1.Controls.Add(Me.Label1)
		Me.GroupBox1.Controls.Add(Me.cleancount)
		Me.GroupBox1.Controls.Add(Me.rawcount)
		Me.GroupBox1.Location = New System.Drawing.Point(827, 12)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(200, 78)
		Me.GroupBox1.TabIndex = 6
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "GroupBox1"
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(48, 48)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(34, 13)
		Me.Label2.TabIndex = 9
		Me.Label2.Text = "Clean"
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(48, 19)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(29, 13)
		Me.Label1.TabIndex = 8
		Me.Label1.Text = "Raw"
		'
		'cleancount
		'
		Me.cleancount.Location = New System.Drawing.Point(94, 45)
		Me.cleancount.Name = "cleancount"
		Me.cleancount.ReadOnly = True
		Me.cleancount.Size = New System.Drawing.Size(100, 20)
		Me.cleancount.TabIndex = 7
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(153, 42)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(132, 23)
		Me.Button1.TabIndex = 7
		Me.Button1.Text = "Save Cleaned Json"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'DataGridView2
		'
		Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
		Me.DataGridView2.Location = New System.Drawing.Point(20, 28)
		Me.DataGridView2.Name = "DataGridView2"
		Me.DataGridView2.ReadOnly = True
		Me.DataGridView2.Size = New System.Drawing.Size(980, 244)
		Me.DataGridView2.TabIndex = 8
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.DataGridView1)
		Me.GroupBox2.Location = New System.Drawing.Point(51, 96)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(1017, 285)
		Me.GroupBox2.TabIndex = 9
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Raw Data"
		'
		'GroupBox3
		'
		Me.GroupBox3.Controls.Add(Me.DataGridView2)
		Me.GroupBox3.Location = New System.Drawing.Point(51, 396)
		Me.GroupBox3.Name = "GroupBox3"
		Me.GroupBox3.Size = New System.Drawing.Size(1017, 291)
		Me.GroupBox3.TabIndex = 10
		Me.GroupBox3.TabStop = False
		Me.GroupBox3.Text = "Clean Data"
		'
		'AFFfix
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1157, 805)
		Me.ControlBox = False
		Me.Controls.Add(Me.GroupBox3)
		Me.Controls.Add(Me.GroupBox2)
		Me.Controls.Add(Me.Button1)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.Clean)
		Me.Controls.Add(Me.jsoninput)
		Me.Controls.Add(Me.jsonfile)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
		Me.Name = "AFFfix"
		Me.Text = "AFFfix"
		Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox1.ResumeLayout(False)
		Me.GroupBox1.PerformLayout()
		CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox3.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub

	Friend WithEvents jsonfile As TextBox
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents jsoninput As Button
	Friend WithEvents DataGridView1 As DataGridView
	Friend WithEvents Clean As Button
    Friend WithEvents rawcount As TextBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents cleancount As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents GroupBox3 As GroupBox
End Class
