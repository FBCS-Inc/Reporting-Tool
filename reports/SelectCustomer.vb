Imports System.Windows.Forms

Public Class SelectCustomer
	Public clist As String = ""
	Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click

		If CustomerIDListBox.SelectedIndices.Count > 0 Then
			For Each xitem In CustomerIDListBox.SelectedItems
				clist += "," + xitem(0).ToString
			Next

			My.Forms.Reporting.customerlist = clist.Substring(1)
		End If

		Me.Close()
	End Sub

	Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
		Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.Close()
	End Sub

	Private Sub SelectCustomer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Try
			FactTableAdapter.ClearBeforeFill = True
			Me.FactTableAdapter.Fill(Me.Collect2000DataSet.Fact)
			Me.CustomCustGroupsTableAdapter.fill(Me.Collect2000DataSet.CustomCustGroups)
		Catch ex As System.Exception
			System.Windows.Forms.MessageBox.Show(ex.Message)
		End Try

	End Sub

	Private Sub NameComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles NameComboBox.SelectedIndexChanged
		FactTableAdapter.ClearBeforeFill = True
		FactTableAdapter.FillByccg(Me.Collect2000DataSet.Fact, NameComboBox.SelectedValue)

	End Sub
End Class
