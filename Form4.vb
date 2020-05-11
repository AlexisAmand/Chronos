Public Class Form4

    Private Sub Form4_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Dim option_astuces As Boolean = True

        Me.Text = "Options"
        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2

        CheckBox1.Text = "Afficher les astuces au démarrage"
        Button1.Text = "Valider"

        FileOpen(1, "config.cfg", OpenMode.Input)

        Input(1, option_astuces)

        FileClose(1)

        If (option_astuces = True) Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If



    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        ' enregistrement des paramétres de l'utilisateur

        ' nombre de checkbox
        ' ce serait bien de faire une boucle for...

        ' booleen qui correspond à l'affichage des astuces au démarrage
        Dim option_astuces As String

        option_astuces = CheckBox1.Checked

        FileOpen(1, "config.cfg", OpenMode.Output)

        Write(1, option_astuces)

        FileClose(1)


        Me.Close()

    End Sub
End Class