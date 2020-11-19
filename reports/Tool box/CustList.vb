Public Class CustList
	Private _clist As String
	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		My.Forms.SelectCustomer.ShowDialog()
		_clist = My.Forms.SelectCustomer.clist
		Me.TextBox1.Text = _clist
	End Sub
	Public ReadOnly Property Clist() As String
		Get
			Return _clist
		End Get
	End Property

End Class
