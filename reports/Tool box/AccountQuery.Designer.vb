<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AccountQuery
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
        Me.AccountBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.QueryButton = New System.Windows.Forms.Button()
        Me.FileButton = New System.Windows.Forms.Button()
        Me.FileTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.excelfiletype = New System.Windows.Forms.RadioButton()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.RptDataGridView = New System.Windows.Forms.DataGridView()
        Me.ExportButton = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.StopWatch = New System.Windows.Forms.Timer(Me.components)
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'AccountBox
        '
        Me.AccountBox.Location = New System.Drawing.Point(253, 62)
        Me.AccountBox.Name = "AccountBox"
        Me.AccountBox.Size = New System.Drawing.Size(202, 20)
        Me.AccountBox.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(160, 65)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Account Number"
        '
        'QueryButton
        '
        Me.QueryButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QueryButton.Location = New System.Drawing.Point(34, 57)
        Me.QueryButton.Name = "QueryButton"
        Me.QueryButton.Size = New System.Drawing.Size(100, 23)
        Me.QueryButton.TabIndex = 73
        Me.QueryButton.Text = "Run Query"
        Me.QueryButton.UseVisualStyleBackColor = True
        '
        'FileButton
        '
        Me.FileButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileButton.Location = New System.Drawing.Point(34, 23)
        Me.FileButton.Name = "FileButton"
        Me.FileButton.Size = New System.Drawing.Size(100, 23)
        Me.FileButton.TabIndex = 77
        Me.FileButton.Text = "Output File"
        Me.FileButton.UseVisualStyleBackColor = True
        '
        'FileTextBox
        '
        Me.FileTextBox.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileTextBox.Location = New System.Drawing.Point(144, 23)
        Me.FileTextBox.Name = "FileTextBox"
        Me.FileTextBox.Size = New System.Drawing.Size(588, 25)
        Me.FileTextBox.TabIndex = 76
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.excelfiletype)
        Me.GroupBox3.Controls.Add(Me.RadioButton1)
        Me.GroupBox3.Location = New System.Drawing.Point(753, 23)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(200, 72)
        Me.GroupBox3.TabIndex = 75
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
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.RptDataGridView)
        Me.GroupBox2.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(34, 133)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(957, 185)
        Me.GroupBox2.TabIndex = 78
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
        Me.RptDataGridView.Size = New System.Drawing.Size(951, 161)
        Me.RptDataGridView.TabIndex = 64
        '
        'ExportButton
        '
        Me.ExportButton.Font = New System.Drawing.Font("Times New Roman", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExportButton.Location = New System.Drawing.Point(34, 89)
        Me.ExportButton.Name = "ExportButton"
        Me.ExportButton.Size = New System.Drawing.Size(100, 23)
        Me.ExportButton.TabIndex = 74
        Me.ExportButton.Text = "Export"
        Me.ExportButton.UseVisualStyleBackColor = True
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
        'AccountQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1008, 450)
        Me.ControlBox = False
        Me.Controls.Add(Me.ExportButton)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.FileButton)
        Me.Controls.Add(Me.FileTextBox)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.QueryButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.AccountBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "AccountQuery"
        Me.Text = "AccountQuery"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        CType(Me.RptDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents AccountBox As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents QueryButton As Button
    Friend WithEvents FileButton As Button
    Friend WithEvents FileTextBox As TextBox
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents excelfiletype As RadioButton
    Friend WithEvents RadioButton1 As RadioButton
    Friend WithEvents GroupBox2 As GroupBox
    Friend WithEvents RptDataGridView As DataGridView
    Friend WithEvents ExportButton As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents StopWatch As Timer
End Class
