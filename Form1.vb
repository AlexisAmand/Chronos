Option Explicit On



Public Class Form1
    Dim njG(8) As String
    Dim njR(11) As String


    Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        Me.Visible = False
        SplashScreen1.visible = true
        SplashScreen1.Update()
        System.Threading.Thread.Sleep(2000)
        SplashScreen1.Close()
        Me.Visible = True



        ' booleen qui correspond à l'affichage des astuces au démarrage
        Dim option_astuces As Boolean = True
        Dim i As Long ' chargement des objets et constantes

        njG(0) = "dimanche"
        njG(1) = "lundi"
        njG(2) = "mardi"
        njG(3) = "mercredi"
        njG(4) = "jeudi"
        njG(5) = "vendredi"
        njG(6) = "samedi"
        njG(7) = "--------"

        '
        ' affichage de la petite boite avec les astuces si l'option est choisie
        '



        ' FileOpen(1, "config.cfg", OpenMode.Input)

        ' Input(1, option_astuces)

        'FileClose(1)

        'If (option_astuces = True) Then
        'Form2_astuces.Visible = True
        'Else
        'Form2_astuces.Visible = False
        'End If

        '
        For i = 1 To 36 ' jour
            If i < 32 Then ComboBox1.Items.Add(Format(i, "00")) ' grégorien
            ComboBox4.Items.Add(Format(i, "00")) ' républicain
        Next i

        ComboBox1.SelectedIndex = 21
        ComboBox4.SelectedIndex = 0
        '
        ComboBox2.Items.Add("janvier") ' mois
        ComboBox2.Items.Add("février")
        ComboBox2.Items.Add("mars")
        ComboBox2.Items.Add("avril")
        ComboBox2.Items.Add("mai")
        ComboBox2.Items.Add("juin")
        ComboBox2.Items.Add("juillet")
        ComboBox2.Items.Add("août")
        ComboBox2.Items.Add("septembre")
        ComboBox2.Items.Add("octobre")
        ComboBox2.Items.Add("novembre")
        ComboBox2.Items.Add("décembre")
        ComboBox2.SelectedIndex = 8
        '
        For i = 1792 To 1805 ' ans G
            ComboBox3.Items.Add(i)
        Next i
        ComboBox3.SelectedIndex = 0
        '
        ' Ajout des jours républicains dans la combobox
        '
        njR(0) = "decadi"
        njR(1) = "primodi"
        njR(2) = "duodi"
        njR(3) = "tridi"
        njR(4) = "quartidi"
        njR(5) = "quintidi"
        njR(6) = "sextidi"
        njR(7) = "septidi"
        njR(8) = "octodi"
        njR(9) = "nonidi"
        njR(10) = "--------"
        '
        ' Ajout des mois républicains dans la combobox 
        '
        ComboBox5.Items.Add("vendémiaire")
        ComboBox5.Items.Add("brumaire")
        ComboBox5.Items.Add("frimaire")
        ComboBox5.Items.Add("nivôse")
        ComboBox5.Items.Add("pluviôse")
        ComboBox5.Items.Add("ventôse")
        ComboBox5.Items.Add("germinal")
        ComboBox5.Items.Add("floréal")
        ComboBox5.Items.Add("prairial")
        ComboBox5.Items.Add("messidor")
        ComboBox5.Items.Add("thermidor")
        ComboBox5.Items.Add("fructidor")

        '
        ' Ajout des ans républicains dans la combobox
        '

        ComboBox6.Items.Add("an I")
        ComboBox6.Items.Add("an II")
        ComboBox6.Items.Add("an III")
        ComboBox6.Items.Add("an IV")
        ComboBox6.Items.Add("an V")
        ComboBox6.Items.Add("an VI")
        ComboBox6.Items.Add("an VII")
        ComboBox6.Items.Add("an VIII")
        ComboBox6.Items.Add("an IX")
        ComboBox6.Items.Add("an X")
        ComboBox6.Items.Add("an XI")
        ComboBox6.Items.Add("an XII")
        ComboBox6.Items.Add("an XIII")
        ComboBox6.Items.Add("an XIV")

        '
        ' Réglages de la fenêtre
        '

        Label1.Text = "Date"
        Label2.Text = "Date"

        Button1.Text = "Date républicaine"
        Button2.Text = "Date grégorienne"

        ToolStripStatusLabel1.Text = "Bienvenue..."
        Me.Text = My.Application.Info.Title
        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2
        Me.Width = 345
        Me.Height = 392


        'MsgBox(Application.StartupPath)

    End Sub


    Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim d1 As Date = 22 & "/" & 9 & "/" & 1792 ' calcule date républicaine
        Dim dG As Date
        Dim jjG As Byte = ComboBox1.SelectedIndex + 1
        Dim mmG As Byte = ComboBox2.SelectedIndex + 1
        Dim aaaaG As Integer = ComboBox3.SelectedIndex + 1792
        Dim jjR As Byte = 1
        Dim mmR As Byte = 1
        Dim aaR As Byte = 1
        Dim i As Integer
        Dim al As String = "I"

        Try
            dG = jjG & "/" & mmG & "/" & aaaaG
        Catch ex As Exception
            GoTo sort
        End Try
        If aaaaG = 1792 Then
            If mmG < 9 Then GoTo sort
            If mmG = 9 And jjG < 22 Then
                'sort:           MsgBox("La date que vous avez indiqué est invalide elle doit être comprise entre le " & vbLf & vbLf & "Du 22 septembre 1792 au 31 décembre 1805", vbExclamation)                
sort:           MsgBox("La date que vous avez indiqué est invalide!" & vbLf & "Elle doit être comprise entre le 22 septembre 1792 et le 31 décembre 1805. Vous devez aussi éviter le 29 février ou le 31 avril.", vbExclamation)
                Label1.Text = njG(7)
                Exit Sub
            End If
        End If
        '
        Dim dateG As New DateTime(aaaaG, mmG, jjG) ' dimanche = 0
        Label1.Text = njG(dateG.DayOfWeek)
        '
        For i = 1 To DateDiff(DateInterval.Day, d1, dG)
            jjR = jjR + 1
            Select Case mmR
                Case 1 To 11 ' mois 30 j
                    If jjR > 30 Then
                        jjR = 1
                        mmR = mmR + 1 ' an 12 mois
                    End If
                Case 12 ' mois 35 ou 36 j
                    Select Case aaR
                        Case 1, 2 : GoTo nb ' 1792-93 1794
                        Case 3 : If jjR > 36 Then GoTo rab ' 1795
                        Case 4, 5, 6 : GoTo nb ' 1796 1797 1798
                        Case 7 : If jjR > 36 Then GoTo rab ' 1799
                        Case 8, 9, 10 : GoTo nb ' 1800, 1801, 1802
                        Case 11 : If jjR > 36 Then GoTo rab ' 1803
                        Case 12, 13, 14 ' 1804 1805
nb:                         If jjR > 35 Then ' mois 36 bissextile
rab:                            jjR = 1
                                mmR = 1
                                aaR = aaR + 1 ' an 1792 - 1805 = 14
                            End If
                    End Select ' aaR
            End Select ' mmR
        Next i
        '
        ComboBox4.SelectedIndex = jjR
        Label2.Text = njR(CInt(Mid(ComboBox4.Text, 2, 1)))
        ComboBox5.SelectedIndex = mmR - 1
        ComboBox6.SelectedIndex = aaR - 1

        '
        ' Les sans-culottides
        '
        ' du 17 au 21 sept -> il s'agit des jours supplémentaires
        ' 
        ' transformation des années 1,2,3,4,5... en années I.II.III 

        Select Case aaR

            Case 1 : al = "I"
            Case 2 : al = "II"
            Case 2 : al = "III"
            Case 2 : al = "IV"
            Case 2 : al = "V"
            Case 2 : al = "VI"
            Case 2 : al = "VII"
            Case 2 : al = "VIII"
            Case 2 : al = "IX"
            Case 2 : al = "X"
            Case 2 : al = "XI"
            Case 2 : al = "XII"
            Case 2 : al = "XIII"
            Case 2 : al = "XIV"
            Case 2 : al = "XV"

        End Select

        If (mmG = 9) Then
            Select Case jjG

                Case 17
                    ToolStripStatusLabel1.Text = " "
                    MsgBox("Il s'agit du 1er jour supplémentaire de l'an " & al & ", il était aussi nommé jour de la vertu")
                    ComboBox4.Text = "--"
                    ComboBox5.Text = "----------------"


                Case 18
                    ToolStripStatusLabel1.Text = " "
                    MsgBox("Il s'agit du 2e jour supplémentaire de l'an " & al & ", il était aussi nommé jour du génie")
                    ComboBox4.Text = "--"
                    ComboBox5.Text = "----------------"

                Case 19
                    ToolStripStatusLabel1.Text = " "
                    MsgBox("Il s'agit du 3e jour supplémentaire de l'an " & al & ", il était aussi nommé jour du travail")
                    ComboBox4.Text = "--"
                    ComboBox5.Text = "----------------"

                Case 20
                    ToolStripStatusLabel1.Text = " "
                    MsgBox("Il s'agit du 4e jour supplémentaire de l'an " & al & ", il était aussi nommé jour du l'opinion")
                    ComboBox4.Text = "--"
                    ComboBox5.Text = "----------------"


                Case 21
                    ToolStripStatusLabel1.Text = " "
                    MsgBox("Il s'agit du 5e jour supplémentaire de l'an " & al & ", il était aussi nommé jour des récompenses")
                    ComboBox4.Text = "--"
                    ComboBox5.Text = "----------------"


            End Select

        End If

        ' étude du cas où il y a le jour sexite
        ' les jours sont donc décalés de 1... logiquement...

        If (aaR + 1 Mod 4 = 0 And mmG = 9 And jjG = 22) Then

            MsgBox("Il s'agit du 6e jour supplémentaire de l'an " & al & ", il était aussi nommé jour de la révolution")

        End If

        If (aaR Mod 4 = 0) Then
            'ComboBox4.Text = jjR + 1
            ComboBox4.Text = jjR
        End If






        'jour de la vertu (17 septembre)
        'jour du génie (18 septembre)
        'jour du travail (19 septembre)
        'jour de l'opinion (20 septembre)
        'jour des récompenses (21 septembre)
        'jour de la révolution (seulement les années sextiles, 22 septembre)






    End Sub


    Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim d1 As Date = 21 & "/" & 9 & "/" & 1792 ' calcule date grégorienne
        Dim ladate As Date
        Dim jjG As Byte
        Dim mmG As Byte
        Dim aaaaG As Integer
        Dim jjRdemande As Byte = ComboBox4.SelectedIndex + 1
        Dim mmRdemande As Byte = ComboBox5.SelectedIndex + 1
        Dim aaRdemande As Byte = ComboBox6.SelectedIndex + 1
        Dim jjR As Byte = 1
        Dim mmR As Byte = 1
        Dim aaR As Byte = 1
        Dim i As Integer
        Dim dMaxi As Date = "31/12/1805"
        '
        For i = 1 To DateDiff(DateInterval.Day, d1, dMaxi) ' 4848
            ladate = d1.AddDays(i)
            If aaRdemande = aaR Then
                If mmRdemande = mmR Then
                    If jjRdemande = jjR Then GoTo ok
                End If
            End If
            jjR = jjR + 1
            Select Case mmR
                Case 1 To 11
                    If jjR > 30 Then
                        jjR = 1
                        mmR = mmR + 1
                    End If
                Case 12
                    Select Case aaR
                        Case 1, 2 : GoTo nb ' 1792-93 1794
                        Case 3 : If jjR > 36 Then GoTo rab ' 1795
                        Case 4, 5, 6 : GoTo nb ' 1796 1797 1798
                        Case 7 : If jjR > 36 Then GoTo rab ' 1799
                        Case 8, 9, 10 : GoTo nb ' 1800, 1801, 1802
                        Case 11 : If jjR > 36 Then GoTo rab ' 1803
                        Case 12, 13, 14 ' 1804 1805
nb:                         If jjR > 35 Then
rab:                            jjR = 1
                                mmR = 1
                                aaR = aaR + 1
                            End If
                    End Select ' aaR
            End Select ' mmR
        Next i
        MsgBox("La date que vous avez indiqué est invalide!" & vbLf & "Elle doit être comprise entre le 1 vendémiaire an I et le 10 nivôse an XIV", vbExclamation)
        'MsgBox("Date invalide : (j/m/a) " & vbLf & vbLf & "Du 1 vendémiaire I au 10 nivôse XIV", vbExclamation)
        Label2.Text = njR(10)
        Exit Sub
        '
ok:     Label2.Text = njR(CInt(Mid(ComboBox4.Text, 2, 1)))
        jjG = CByte(Mid(ladate, 1, 2)) - 1
        mmG = CByte(Mid(ladate, 4, 2)) - 1
        aaaaG = CInt(Mid(ladate, 7, 4)) - 1792
        ComboBox1.SelectedIndex = jjG
        Label1.Text = njG(ladate.DayOfWeek)
        ComboBox2.SelectedIndex = mmG
        ComboBox3.SelectedIndex = aaaaG
    End Sub

    Private Sub QuitterToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If MsgBox("Etes-vous sûr de vouloir quitter ?", vbYesNo) = vbYes Then
            Me.Close()
        End If

    End Sub

    Private Sub AProposToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub Button1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseEnter
        ToolStripStatusLabel1.Text = "Convertir en date républicaine"
    End Sub


    Private Sub Button1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub Button2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseEnter
        ToolStripStatusLabel1.Text = "Convertir en date grégorienne"
    End Sub

    Private Sub Button2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.MouseLeave
        ToolStripStatusLabel1.Text = " "

    End Sub

    Private Sub ComboBox1_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le jour de la date grégorienne"
    End Sub

    Private Sub ComboBox2_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le mois de la date grégorienne"
    End Sub

    Private Sub ComboBox3_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez l'année de la date grégorienne"
    End Sub

    Private Sub ComboBox1_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox2_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox3_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox4_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le jour de la date républicaine"
    End Sub

    Private Sub ComboBox5_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez le mois de la date républicaine"
    End Sub

    Private Sub ComboBox6_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.MouseEnter
        ToolStripStatusLabel1.Text = "Choisissez l'année de la date républicaine"
    End Sub

    Private Sub ComboBox4_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox4.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox5_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox5.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub ComboBox6_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox6.MouseLeave
        ToolStripStatusLabel1.Text = " "
    End Sub

  




    Private Sub SiteWebDuLogicielToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub QuitterToolStripMenuItem_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Application.Exit()


    End Sub

    Private Sub QuitterToolStripMenuItem_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = "Quitter le logiciel"
    End Sub


    Private Sub QuitterToolStripMenuItem_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = " "
    End Sub


    Private Sub SiteWebDuLogicielToolStripMenuItem_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = "Visiter le site web du logiciel"
    End Sub

    Private Sub SiteWebDuLogicielToolStripMenuItem_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = " "

    End Sub


    Private Sub AProposToolStripMenuItem_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = "A propos de Chronos"
    End Sub


    Private Sub AProposToolStripMenuItem_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        ToolStripStatusLabel1.Text = " "
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub GrégorienEtRépublicainToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub AnnéeBissextilToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MenuStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles MenuStrip1.ItemClicked

    End Sub

    Private Sub TrouverLeJourDuneDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MajToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)


    End Sub

    Private Sub VérifierLesMisesÀJourToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)








    End Sub


    Private Sub TrouverLeJourDuneDateToolStripMenuItem_MouseEnter(sender As Object, e As System.EventArgs)

        ToolStripStatusLabel1.Text = "A quel jour correspond une date ?"

    End Sub



    Private Sub TrouverLeJourDuneDateToolStripMenuItem_MouseLeave(sender As Object, e As System.EventArgs)
        ToolStripStatusLabel1.Text = ""
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        Form4.Visible = True

    End Sub

    Private Sub OptionsToolStripMenuItem_MouseEnter(sender As Object, e As System.EventArgs)

        ToolStripStatusLabel1.Text = "Options du logiciel"

    End Sub


    Private Sub MenuToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MenuToolStripMenuItem.Click

    End Sub

    Private Sub QuitterToolStripMenuItem_Click_2(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub TestToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TestToolStripMenuItem.Click
        System.Diagnostics.Process.Start("http://chronos.genealexis.fr")
    End Sub

    Private Sub TrouverLeJourDuneDateToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles TrouverLeJourDuneDateToolStripMenuItem1.Click
        Form3.Visible = True
    End Sub

    Private Sub AnnéeBissextileToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AnnéeBissextileToolStripMenuItem.Click

        Form2.Visible = True

    End Sub

    Private Sub AProposToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles AProposToolStripMenuItem1.Click
        AboutBox1.Show()
    End Sub

    Private Sub QuitterToolStripMenuItem_Click_3(sender As System.Object, e As System.EventArgs) Handles QuitterToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Label1_Click(sender As System.Object, e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub CalculDuJourJulienToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CalculDuJourJulienToolStripMenuItem.Click

        jourjulien.Show()

    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub CopierLeResultatToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub EncyclopédieToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TestToolStripMenuItem1_Click(sender As Object, e As EventArgs)

    End Sub
End Class
