Public Class Form1

    Dim TCP_Client1 As New TCP_Class1

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        TCP_Client1.Connect()

    End Sub

    Private Sub Send_Request()


    End Sub


    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        TCP_Client1.Send()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        TCP_Client1.close()
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As System.EventArgs) Handles ListBox1.DoubleClick
        Me.ListBox1.Items.Clear()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

    End Sub
End Class
