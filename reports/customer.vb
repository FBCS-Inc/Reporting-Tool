Public Class Customer


	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		My.Forms.SelectCustomer.ShowDialog()

		TextBox1.Text = My.Forms.SelectCustomer.clist
	End Sub
End Class