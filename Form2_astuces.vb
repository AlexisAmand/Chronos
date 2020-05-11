Public Class Form2_astuces

    'numéro de l'actuce à afficher
    Dim X As Integer
    ' tableau qui contient les textes des astuces a afficher
    Dim astuces(50) As String
    ' nombre d'astuces affichables
    Dim nb_tot_astuces As Integer = 16

    Private Sub Form2_astuces_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        '

        Label1.AutoSize = True
        Label1.TextAlign = ContentAlignment.MiddleCenter

        '
        ' Création du tableau qui contient les astuces
        '

        astuces(1) = "L'expression calendrier romain désigne l'ensemble des calendriers utilisés par les Romains jusqu'à la création du calendrier julien en 45 av. J.-C. Il serait, selon la tradition entre autres rapportée par Ovide dans Les Fastes, l'invention de Romulus, fondateur de Rome vers 753 av. J.-C.. Il semble cependant avoir été fondé sur le calendrier lunaire grec ou étrusque."

        astuces(2) = "Le 22 fructidor an XIII (9 septembre 1805), Napoléon signa le sénatus-consulte qui abrogea le calendrier républicain et instaura le retour au calendrier grégorien à partir du 1er janvier 1806."

        astuces(3) = "Le mois vendémiaire tire son nom des vendanges qui ont lieu de septembre en octobre."

        astuces(4) = "Le mois brumaire tire son nom des brouillards et des brumes basses qui sont la transudation de la nature d'octobre en novembre."

        astuces(5) = "Le mois frimaire tire son nom du froid, tantôt sec, tantôt humide, qui se fait sentir de novembre en décembre."

        astuces(6) = "Le mois nivôse tire son nom de la neige qui blanchit la terre de décembre en janvier."

        astuces(7) = "Le mois pluviôse tire son nom des pluies qui tombent généralement avec plus d'abondance de janvier en février."

        astuces(8) = "Le mois ventôse tire son nom des giboulées qui ont lieu, et du vent qui vient sécher la terre de février en mars."

        astuces(9) = "Le mois germinal tire son nom de la fermentation et du développement de la sève de mars en avril."

        astuces(10) = "Le mois floréal tire son nom de l'épanouissement des fleurs d'avril en mai."

        astuces(11) = "Le mois prairial tire son nom de la fécondité riante et de la récolte des prairies de mai en juin."

        astuces(12) = "Le mois messidor tire son nom de l'aspect des épis ondoyants et des moissons dorées qui couvrent les champs de juin en juillet."

        astuces(13) = "Le mois thermidor tire son nom de la chaleur tout à la fois solaire et terrestre qui embrase l'air de juillet en août."

        astuces(14) = "Le mois fructidor tire son nom des fruits que le soleil dore & mûrit d'août en septembre."

        astuces(15) = "le mois de brumaire a donné son nom au coup d'État, intervenu le 18 brumaire an VIII (le 9 novembre 1799), qui porta le général Napoléon Bonaparte au pouvoir en France. On parle dans ce cas du 18-Brumaire (nom propre, avec trait d'union et majuscule à la deuxième partie du nom), mais également de Brumaire (là aussi avec une majuscule, pour distinguer du nom du mois en général)."

        astuces(16) = "Le décret du 4 frimaire an II (24 novembre 1793) « sur l'ère, le commencement et l'organisation de l'année, et sur les noms des jours et des mois » orthographiait le nom du mois nivose, sans accent circonflexe. L'ajout généralisé de cet accent s'est installé progressivement, à une époque ultérieure indéterminée. On rencontre d'ailleurs des milliers d'actes ou documents officiels de l'époque ne faisant pas encore usage de cet accent."

        ' option de la fenêtre

        Me.Text = "Le saviez vous ?"
        Me.Left = (SystemInformation.PrimaryMonitorSize.Width - Me.Width) / 2
        Me.Top = (SystemInformation.PrimaryMonitorSize.Height - Me.Height) / 2

        'la taille de la fenêtre
        Me.Width = 370
        Me.Height = 200

        'la fenêtre au premier plan
        Me.TopMost = True

        ' le texte des boutons
        Button1.Text = "Précédente"
        Button2.Text = "Suivante"
        Button3.Text = "Fermer"

        ' Initialize the random-number generator.
        Randomize()

        ' Generate random value between 1 and nb_tot_astuces.
        X = CInt(Int((nb_tot_astuces * Rnd()) + 1))
        Label1.Text = "Astuce " & X & "/" & nb_tot_astuces
        RichTextBox1.Text = astuces(X)




    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        X = X + 1
        If X = nb_tot_astuces + 1 Then
            X = 1
        End If
        Label1.Text = "Astuce " & X & "/" & nb_tot_astuces
        RichTextBox1.Text = astuces(X)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        X = X - 1
        If X = 0 Then
            X = nb_tot_astuces
        End If
        Label1.Text = "Astuce " & X & "/" & nb_tot_astuces
        RichTextBox1.Text = astuces(X)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Visible = False

    End Sub
End Class