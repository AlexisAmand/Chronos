Public Class Form2

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Bis As Integer


        'la taille de la fenêtre
        Me.Width = 370
        Me.Height = 126

        'la fenêtre au premier plan
        Me.TopMost = True

        If (IsNumeric(TextBox1.Text)) Then

            Bis = TextBox1.Text

            'soit divisibles par 4 mais non divisibles par 100
            'soit divisibles par 400.
            If (Bis Mod 4 = 0 And Bis Mod 100 <> 0) Or (Bis Mod 400 = 0) Then
                MsgBox("C'est une année bissextile!")
            Else
                MsgBox("Ce n'est pas une année bissextile!")
            End If

        Else

            MsgBox("Vous n'avez pas indiqué une année correcte")

        End If









    End Sub

    Private Sub Form2_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.Text = "Année bissextile ?"
        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub
End Class
