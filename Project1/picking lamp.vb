Public Class picking_lamp
    Public liffting As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not TextBox1.Text = "" Then
            Return
        Else
            liffting = TextBox1.Text
            Me.Hide()
            Form2.Show()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class