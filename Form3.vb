Public Class Form3

    ' Cette form trouve a quel jour correspond une date donnée

    Dim jour As Integer
    Dim fin_an As Integer
    Dim debut_an As Integer
    Dim mois, temp, Bis As Integer




    Private Sub Form3_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.Text = "Trouver le jour d'une date"

        Dim i, j, k As Integer

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

        For j = 17 To 21
            ComboBox3.Items.Add(Format(j, "00"))
        Next j

        For k = 0 To 99
            ComboBox4.Items.Add(Format(k, "00"))
        Next k

        ComboBox1.SelectedIndex = 2
        ComboBox2.SelectedIndex = 2
        ComboBox3.SelectedIndex = 2
        ComboBox4.SelectedIndex = 2

        Button1.Text = "Calculer"

        Me.Text = "Jour d'une date"
        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        ' Voici le mois qui a été selectionné dans la Combobox 2
        mois = Val(ComboBox2.SelectedIndex) + 1

        'Voici le jour qui a été selectionné dans la Combobox 1
        jour = Val(ComboBox1.SelectedIndex) + 1

        'Voici le siécle qui a été selectionné dans la Combobox 3
        debut_an = Val(ComboBox3.SelectedIndex) + 17

        'Voici les 2 derniers chiffres de l'année dans la Combobox 4
        fin_an = Val(ComboBox4.SelectedIndex)

        'MsgBox("la date choisie est le " & jour & "/" & mois & "/" & debut_an & fin_an)

        ' En Italique, calcul du jour de la semaine de l'exemple 1er août 1947.
        ' On garde les deux derniers chiffres de l'année en question; (1947 => 47)
        ' On ajoute 1/4 de ce chiffre en ignorant les restes; (47/4 = 11, reste 3 ignoré)
        ' On ajoute la journée du mois; (dans ce cas => 1)

        temp = fin_an + (fin_an \ 4) + jour

        ' Selon le mois on ajoute : (Août => 3)
        ' Janvier = 1
        ' Février = 4
        ' Mars = 4
        ' Avril = 0
        ' Mai = 2
        ' Juin = 5
        ' Juillet = 0
        ' Août = 3
        ' Septembre = 6
        ' Octobre = 1
        ' Novembre = 4
        ' Décembre = 6

        Select Case mois
            Case "1", "10"
                temp = temp + 1
            Case "2", "3", "11"
                temp = temp + 4
            Case "4", "7"
                temp = temp + 0
            Case "5"
                temp = temp + 2
            Case "6"
                temp = temp + 5
            Case "8"
                temp = temp + 3
            Case "9", "12"
                temp = temp + 6
        End Select

        ' Si l'année est bissextile et le mois est janvier ou février, on ôte 1 , (1947 =>année non bissextile)

        Bis = debut_an * 100 + fin_an

        If (Bis Mod 4 = 0 And Bis Mod 100 <> 0) Or (Bis Mod 400 = 0) Then
            temp = temp - 1
        End If

        ' Selon le siècle, on ajoute : (19** => 0)
        ' Années 1700 = 4;
        ' Années 1800 = 2;
        ' Années 1900 = 0;
        ' Années 2000 = 6;
        ' Années 2100 = 4;

        Select Case debut_an
            Case "17", "21"
                temp = temp + 4
            Case "18"
                temp = temp + 2
            Case "19"
                temp = temp + 0
            Case "20"
                temp = temp + 6
        End Select

        temp = temp Mod 7

        ' MsgBox("le jour est " & temp)

        ' On divise la somme par 7 et on garde le reste; (47 + 11 + 1 + 3 - 0 + 0 = 62; 62 divisé par 7 = 8, reste 6)
        ' Le reste représente le jour de la semaine recherché: (Le 1er août 1947 était un vendredi)
        ' 1 pour Dimanche,
        ' 2 pour Lundi,
        ' 3 pour Mardi,
        ' 4 pour Mercredi,
        ' 5 pour Jeudi,
        ' 6 pour Vendredi,
        ' 0 pour Samedi,

        Select Case temp
            Case "1"
                MsgBox("C'était un lundi")
            Case "2"
                MsgBox("C'était un mardi")
            Case "3"
                MsgBox("C'était un mercredi")
            Case "4"
                MsgBox("C'était un jeudi")
            Case "5"
                MsgBox("C'était un vendredi")
            Case "6"
                MsgBox("C'était un samedi")
            Case "0"
                MsgBox("C'était un dimanche")
        End Select


    End Sub

    Private Sub ComboBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le jour"



    End Sub

    Private Sub ComboBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub ComboBox2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le mois"
    End Sub

    Private Sub ComboBox2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub

    Private Sub ComboBox3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez les deux premiers chiffres de l'année"
    End Sub

    Private Sub ComboBox4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez les deux derniers chiffres de l'année"
    End Sub

    Private Sub ComboBox4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub Button1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseEnter
        ToolStripStatusLabel1.Text = "Cliquez pour trouver le jour"
    End Sub

    Private Sub Button1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub
End Class