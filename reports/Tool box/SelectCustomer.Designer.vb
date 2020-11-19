<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class SelectCustomer
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
		Me.components = New System.ComponentModel.Container()
		Dim CustomerIDLabel As System.Windows.Forms.Label
		Dim NameLabel As System.Windows.Forms.Label
		Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
		Me.OK_Button = New System.Windows.Forms.Button()
		Me.Cancel_Button = New System.Windows.Forms.Button()
		Me.CustomerIDListBox = New System.Windows.Forms.ListBox()
		Me.NameComboBox = New System.Windows.Forms.ComboBox()
		Me.Collect2000DataSet = New reports.Collect2000DataSet()
		Me.CustomCustGroupsBindingSource = New System.Windows.Forms.BindingSource(Me.components)
		Me.CustomCustGroupsTableAdapter = New reports.collect2000DataSetTableAdapters.CustomCustGroupsTableAdapter()
		Me.TableAdapterManager = New reports.collect2000DataSetTableAdapters.TableAdapterManager()
		Me.FactBindingSource = New System.Windows.Forms.BindingSource(Me.components)
		Me.FactTableAdapter = New reports.collect2000DataSetTableAdapters.FactTableAdapter()
		CustomerIDLabel = New System.Windows.Forms.Label()
		NameLabel = New System.Windows.Forms.Label()
		Me.TableLayoutPanel1.SuspendLayout()
		CType(Me.Collect2000DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.CustomCustGroupsBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
		CType(Me.FactBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'CustomerIDLabel
		'
		CustomerIDLabel.AutoSize = True
		CustomerIDLabel.Location = New System.Drawing.Point(230, 11)
		CustomerIDLabel.Name = "CustomerIDLabel"
		CustomerIDLabel.Size = New System.Drawing.Size(68, 13)
		CustomerIDLabel.TabIndex = 7
		CustomerIDLabel.Text = "Customer ID:"
		'
		'NameLabel
		'
		NameLabel.AutoSize = True
		NameLabel.Location = New System.Drawing.Point(16, 11)
		NameLabel.Name = "NameLabel"
		NameLabel.Size = New System.Drawing.Size(86, 13)
		NameLabel.TabIndex = 5
		NameLabel.Text = "Customer Group:"
		'
		'TableLayoutPanel1
		'
		Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.TableLayoutPanel1.ColumnCount = 2
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Controls.Add(Me.OK_Button, 0, 0)
		Me.TableLayoutPanel1.Controls.Add(Me.Cancel_Button, 1, 0)
		Me.TableLayoutPanel1.Location = New System.Drawing.Point(277, 274)
		Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
		Me.TableLayoutPanel1.RowCount = 1
		Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
		Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
		Me.TableLayoutPanel1.TabIndex = 0
		'
		'OK_Button
		'
		Me.OK_Button.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.OK_Button.Location = New System.Drawing.Point(3, 3)
		Me.OK_Button.Name = "OK_Button"
		Me.OK_Button.Size = New System.Drawing.Size(67, 23)
		Me.OK_Button.TabIndex = 0
		Me.OK_Button.Text = "OK"
		'
		'Cancel_Button
		'
		Me.Cancel_Button.Anchor = System.Windows.Forms.AnchorStyles.None
		Me.Cancel_Button.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.Cancel_Button.Location = New System.Drawing.Point(76, 3)
		Me.Cancel_Button.Name = "Cancel_Button"
		Me.Cancel_Button.Size = New System.Drawing.Size(67, 23)
		Me.Cancel_Button.TabIndex = 1
		Me.Cancel_Button.Text = "Cancel"
		'
		'CustomerIDListBox
		'
		Me.CustomerIDListBox.DataSource = Me.FactBindingSource
		Me.CustomerIDListBox.DisplayMember = "CustomerID"
		Me.CustomerIDListBox.FormattingEnabled = True
		Me.CustomerIDListBox.Location = New System.Drawing.Point(227, 27)
		Me.CustomerIDListBox.Name = "CustomerIDListBox"
		Me.CustomerIDListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple
		Me.CustomerIDListBox.Size = New System.Drawing.Size(120, 199)
		Me.CustomerIDListBox.TabIndex = 8
		Me.CustomerIDListBox.ValueMember = "CustomerID"
		'
		'NameComboBox
		'
		Me.NameComboBox.DataSource = Me.CustomCustGroupsBindingSource
		Me.NameComboBox.DisplayMember = "Name"
		Me.NameComboBox.FormattingEnabled = True
		Me.NameComboBox.Location = New System.Drawing.Point(12, 27)
		Me.NameComboBox.Name = "NameComboBox"
		Me.NameComboBox.Size = New System.Drawing.Size(180, 21)
		Me.NameComboBox.TabIndex = 6
		Me.NameComboBox.ValueMember = "id"
		'
		'Collect2000DataSet
		'
		Me.Collect2000DataSet.DataSetName = "collect2000DataSet"
		Me.Collect2000DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
		'
		'CustomCustGroupsBindingSource
		'
		Me.CustomCustGroupsBindingSource.DataMember = "CustomCustGroups"
		Me.CustomCustGroupsBindingSource.DataSource = Me.Collect2000DataSet
		'
		'CustomCustGroupsTableAdapter
		'
		Me.CustomCustGroupsTableAdapter.ClearBeforeFill = True
		'
		'TableAdapterManager
		'
		Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
		Me.TableAdapterManager.CustomCustGroupsTableAdapter = Me.CustomCustGroupsTableAdapter
		Me.TableAdapterManager.FactTableAdapter = Nothing
		Me.TableAdapterManager.UpdateOrder = reports.collect2000DataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
		'
		'FactBindingSource
		'
		Me.FactBindingSource.DataMember = "Fact"
		Me.FactBindingSource.DataSource = Me.Collect2000DataSet
		'
		'FactTableAdapter
		'
		Me.FactTableAdapter.ClearBeforeFill = True
		'
		'SelectCustomer
		'
		Me.AcceptButton = Me.OK_Button
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.CancelButton = Me.Cancel_Button
		Me.ClientSize = New System.Drawing.Size(435, 315)
		Me.Controls.Add(CustomerIDLabel)
		Me.Controls.Add(Me.CustomerIDListBox)
		Me.Controls.Add(NameLabel)
		Me.Controls.Add(Me.NameComboBox)
		Me.Controls.Add(Me.TableLayoutPanel1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "SelectCustomer"
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
		Me.Text = "Select Customer"
		Me.TableLayoutPanel1.ResumeLayout(False)
		CType(Me.Collect2000DataSet, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.CustomCustGroupsBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
		CType(Me.FactBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
	Friend WithEvents OK_Button As System.Windows.Forms.Button
	Friend WithEvents Cancel_Button As System.Windows.Forms.Button
	Friend WithEvents CustomerIDListBox As ListBox
	Friend WithEvents NameComboBox As ComboBox
	Friend WithEvents Collect2000DataSet As Collect2000DataSet
	Friend WithEvents CustomCustGroupsBindingSource As BindingSource
	Friend WithEvents CustomCustGroupsTableAdapter As collect2000DataSetTableAdapters.CustomCustGroupsTableAdapter
	Friend WithEvents TableAdapterManager As collect2000DataSetTableAdapters.TableAdapterManager
	Friend WithEvents FactBindingSource As BindingSource
	Friend WithEvents FactTableAdapter As collect2000DataSetTableAdapters.FactTableAdapter
End Class
