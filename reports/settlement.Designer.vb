<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Settlement
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
        Me.ReportData1 = New reports.ReportData()
        Me.SuspendLayout()
        '
        'ReportData1
        '
        Me.ReportData1.Location = New System.Drawing.Point(12, 12)
        Me.ReportData1.Name = "ReportData1"
        Me.ReportData1.Size = New System.Drawing.Size(1564, 495)
        Me.ReportData1.TabIndex = 0
        '
        'settlement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1568, 563)
        Me.ControlBox = False
        Me.Controls.Add(Me.ReportData1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "settlement"
        Me.ShowIcon = False
        Me.Text = "NewRpt"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ReportData1 As ReportData
End Class
