Imports System.Windows.Forms

Public Class DataChangeWarning

	Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
		Me.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.Close()
	End Sub

	Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
		Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.Close()
	End Sub

	Private Sub DataChangeWarning_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		Me.BackColor = GetBackcolor()

	End Sub
	Private Function GetBackcolor() As Color
		Dim c(5) As Color
		c(0) = Color.AliceBlue
		c(1) = Color.LightCoral
		c(2) = Color.Cyan
		c(3) = Color.GreenYellow
		c(4) = Color.LightSalmon
		Randomize()
		' Generate random value between 1 and 6.
		Dim idx As Integer = CInt(Int((4 * Rnd())))
		Return c(idx)
	End Function
End Class
