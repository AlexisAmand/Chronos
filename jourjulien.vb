Public Class jourjulien

    ' Calcul du jour julien

    Private Sub jourjulien_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.Text = "Calcul du jour julien"

        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2

        ' remplissage des combobox

        For i = 1 To 31
            ComboBox1.Items.Add(Format(i, "00"))
        Next i

        ComboBox2.SelectedItem = 4

        ComboBox2.Items.Add("Janvier")
        ComboBox2.Items.Add("Février")
        ComboBox2.Items.Add("Mars")
        ComboBox2.Items.Add("Avril")
        ComboBox2.Items.Add("Mai")
        ComboBox2.Items.Add("Juin")
        ComboBox2.Items.Add("Juillet")
        ComboBox2.Items.Add("Août")
        ComboBox2.Items.Add("Septembre")
        ComboBox2.Items.Add("Octobre")
        ComboBox2.Items.Add("Novembre")
        ComboBox2.Items.Add("Decembre")

        For j = 15 To 21
            ComboBox3.Items.Add(Format(j, "00"))
        Next j

        For k = 0 To 99
            ComboBox4.Items.Add(Format(k, "00"))
        Next k

        ComboBox1.SelectedIndex = 2
        ComboBox2.SelectedIndex = 2
        ComboBox3.SelectedIndex = 2
        ComboBox4.SelectedIndex = 2

        Label1.Text = "Indiquez la date"

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim mois, jour, debut_an, fin_an, annee As Integer
        Dim s, b, jj As Single
        Dim msg, style, title As String

        Dim rep As MsgBoxResult 'cas ou la date est < à 1582
        msg = "Le calcul de jour julien est disponible uniquement pour les dates supérieures à 1582 (année de création du calendrier julien)"
        style = MsgBoxStyle.Information
        title = "Information"


        ' Voici le mois qui a été selectionné dans la Combobox 2
        mois = Val(ComboBox2.SelectedIndex) + 1

        'Voici le jour qui a été selectionné dans la Combobox 1
        jour = Val(ComboBox1.SelectedIndex) + 1

        'Voici le siécle qui a été selectionné dans la Combobox 3
        debut_an = Val(ComboBox3.SelectedIndex) + 15

        'Voici les 2 derniers chiffres de l'année dans la Combobox 4
        fin_an = Val(ComboBox4.SelectedIndex)

        annee = debut_an * 100 + fin_an

        MsgBox("la date choisie est le " & jour & "/" & mois & "/" & annee)

        If (annee > 1582) Then

            If (mois = 1) Or (mois = 2) Then

                annee = annee - 1
                mois = mois + 12

            End If

            s = Int(annee / 100)
            b = 2 - s + Int(s / 4)

            jj = Int(365.25 * (annee + 4716)) + Int(30.6001 * (mois + 1)) + jour + b - 1524.5

        Else

            rep = MsgBox(msg, style, title)

        End If





    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
End Class