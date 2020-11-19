<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataChangeWarning
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
		Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
		Me.OK_Button = New System.Windows.Forms.Button()
		Me.Cancel_Button = New System.Windows.Forms.Button()
		Me.Infolbl = New System.Windows.Forms.Label()
		Me.TableLayoutPanel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'TableLayoutPanel1
		'
		Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.TableLayoutPanel1.ColumnCount = 2
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
		Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
		Me.TableLayoutPanel1.Location = New System.Drawing.Point(508, 506)
		Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
		Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
		Me.TableLayoutPanel1.RowCount = 1
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Size = New System.Drawing.Size(268, 54)
		Me.TableLayoutPanel1.TabIndex = 0
		'
		'OK_Button
		'
		Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.OK_Button.Location = New System.Drawing.Point(6, 6)
		Me.OK_Button.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
		Me.OK_Button.Name = "OK_Button"
		Me.OK_Button.Size = New System.Drawing.Size(122, 42)
		Me.OK_Button.TabIndex = 0
		Me.OK_Button.Text = "OK"
		'
		'Cancel_Button
		'
		Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.Cancel_Button.Location = New System.Drawing.Point(140, 6)
		Me.Cancel_Button.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
		Me.Cancel_Button.Name = "Cancel_Button"
		Me.Cancel_Button.Size = New System.Drawing.Size(122, 42)
		Me.Cancel_Button.TabIndex = 1
		Me.Cancel_Button.Text = "Cancel"
		'
		'Infolbl
		'
		Me.Infolbl.AutoSize = True
		Me.Infolbl.Location = New System.Drawing.Point(24, 24)
		Me.Infolbl.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
		Me.Infolbl.Name = "Infolbl"
		Me.Infolbl.Size = New System.Drawing.Size(66, 24)
		Me.Infolbl.TabIndex = 1
		Me.Infolbl.Text = "Label1"
		'
		'DataChangeWarning
		'
		Me.AcceptButton = Me.OK_Button
		Me.AutoScaleDimensions = New System.Drawing.SizeF(11.0!, 24.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.CancelButton = Me.Cancel_Button
		Me.ClientSize = New System.Drawing.Size(798, 582)
		Me.Controls.Add(Me.Infolbl)
		Me.Controls.Add(Me.TableLayoutPanel1)
		Me.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Margin = New System.Windows.Forms.Padding(6, 6, 6, 6)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "DataChangeWarning"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Data Change "
		Me.TopMost = True
		Me.TableLayoutPanel1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
	Friend WithEvents OK_Button As System.Windows.Forms.Button
	Friend WithEvents Cancel_Button As System.Windows.Forms.Button
	Public WithEvents Infolbl As Label
End Class
