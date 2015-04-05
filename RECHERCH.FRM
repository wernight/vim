VERSION 5.00
Begin VB.Form frmRecherche 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recherche d'un verbe irrégulier"
   ClientHeight    =   2865
   ClientLeft      =   1710
   ClientTop       =   2490
   ClientWidth     =   7800
   Icon            =   "Recherch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2865
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picListLbl 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      Picture         =   "Recherch.frx":030A
      ScaleHeight     =   285
      ScaleWidth      =   7800
      TabIndex        =   9
      Top             =   1060
      Width           =   7800
   End
   Begin VB.ListBox lstRecherche 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      ItemData        =   "Recherch.frx":08DA
      Left            =   0
      List            =   "Recherch.frx":08DC
      TabIndex        =   10
      Top             =   1320
      Width           =   7800
   End
   Begin VB.Frame Frame 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.OptionButton optElément 
         Caption         =   "Participe passé"
         Height          =   195
         Index           =   2
         Left            =   3720
         TabIndex        =   7
         Top             =   650
         Width           =   1575
      End
      Begin VB.OptionButton optElément 
         Caption         =   "Prétérit"
         Height          =   195
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   650
         Width           =   1095
      End
      Begin VB.OptionButton optElément 
         Caption         =   "Infinitif"
         Height          =   195
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   650
         Width           =   1095
      End
      Begin VB.OptionButton optElément 
         Caption         =   "Traduction"
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   8
         Top             =   650
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtRecherche 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton cmdRechercher 
         Caption         =   "Rec&hercher maintenant"
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Top             =   230
         Width           =   2055
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dans:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   650
         Width           =   510
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rechercher:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   285
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmRecherche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lstRechercheID() As Integer

Private Sub cmdRechercher_Click()
Dim Verbe As String, Mot2 As String * 17, VrbID As Integer
Dim Elément As Integer, Recherche As String
    ' Centre
    If Me.Height <> lstRecherche.Top + lstRecherche.Height + 375 Then
        Me.Height = lstRecherche.Top + lstRecherche.Height + 375
        Me.Move (Screen.Width - Me.Width) / 2, _
            (Screen.Height - Me.Height) / 2
    End If
    
    ' Efface
    lstRecherche.Clear
    
    ' Donne l'élément choisi et le texte
    Recherche = UCase(Trim(txtRecherche))
    For i% = 0 To 3
        If optElément(i%) Then Elément = i%
    Next
    
    ' Recherche
    For i% = 1 To NbrVrb
        'Test
        If UCase(Trim(Mid(Vrb(i%), Elément * 30 + 1, 15))) = Recherche Or Recherche = "*" Then
            ' N'affiche que les 15-1er charac.
            Verbe = ""
            For j% = 0 To 2
                Mot2 = Mid(Vrb(i%), j% * 30 + 1, 30)
                Verbe = Verbe & Mot2 & " "
            Next
            Verbe = Verbe & " " & Mid(Vrb(i%), j% * 30 + 1, 30)
            lstRecherche.AddItem Verbe
            ReDim lstRechercheID(lstRecherche.ListCount - 1) As Integer
            lstRechercheID(lstRecherche.ListCount - 1) = i%
        End If
    Next
    
    ' Affiche si rien trouvé
    If lstRecherche.ListCount = 0 Then
        lstRecherche.AddItem "Aucun élément à afficher dans cet aperçu."
        lstRecherche.Enabled = False
    Else
        lstRecherche.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    ' Centre
    Me.Height = Frame.Top + Frame.Height + 406
End Sub

Private Sub lstRecherche_DblClick()
    ' Donne le verbe
    VrbEnCours = lstRechercheID(lstRecherche.ListIndex)
    Hide
End Sub

Private Sub lstRecherche_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then lstRecherche_DblClick
End Sub

Private Sub optElément_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then cmdRechercher_Click
End Sub

Private Sub txtRecherche_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then cmdRechercher_Click
End Sub
