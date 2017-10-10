Public Class Form1


    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Me.WebBrowser1.Navigate(TextBox1.Text)
    End Sub

End Class