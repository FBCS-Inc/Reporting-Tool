<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Customer
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
		Me.CustList1 = New reports.CustList()
		Me.SuspendLayout()
		'
		'CustList1
		'
		Me.CustList1.Location = New System.Drawing.Point(40, 63)
		Me.CustList1.Name = "CustList1"
		Me.CustList1.Size = New System.Drawing.Size(602, 55)
		Me.CustList1.TabIndex = 0
		'
		'Customer
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(800, 450)
		Me.Controls.Add(Me.CustList1)
		Me.Name = "Customer"
		Me.Text = "customer"
		Me.ResumeLayout(False)

	End Sub

	Friend WithEvents CustList1 As CustList
End Class
