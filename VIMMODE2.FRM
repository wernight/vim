VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInterr2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interrogation"
   ClientHeight    =   1740
   ClientLeft      =   555
   ClientTop       =   2910
   ClientWidth     =   9240
   Icon            =   "VIMMode2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1740
   ScaleWidth      =   9240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRéponce 
      Caption         =   "&Donner la réponse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Texte 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   3
      Left            =   6720
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Texte 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   2
      Left            =   4440
      MaxLength       =   30
      TabIndex        =   6
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Texte 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   2280
      MaxLength       =   30
      TabIndex        =   5
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox Texte 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      MaxLength       =   30
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin Threed.SSCommand SScmdVFr 
      Height          =   405
      Left            =   3720
      TabIndex        =   11
      Top             =   1185
      Width           =   3885
      _Version        =   65536
      _ExtentX        =   6853
      _ExtentY        =   714
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      Outline         =   0   'False
      AutoSize        =   2
      Picture         =   "VIMMode2.frx":030A
   End
   Begin Threed.SSCommand SScmdQuitter 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   7790
      TabIndex        =   10
      Top             =   1190
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   714
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      Outline         =   0   'False
      AutoSize        =   2
      Picture         =   "VIMMode2.frx":0E76
   End
   Begin Threed.SSCommand SScmdOK 
      Height          =   405
      Left            =   120
      TabIndex        =   9
      Top             =   1185
      Width           =   1365
      _Version        =   65536
      _ExtentX        =   2408
      _ExtentY        =   714
      _StockProps     =   78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      Outline         =   0   'False
      AutoSize        =   2
      Picture         =   "VIMMode2.frx":15B6
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prétérit"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Participe passé"
      Height          =   195
      Index           =   2
      Left            =   4440
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Traduction"
      Height          =   195
      Index           =   3
      Left            =   6720
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Infinitif"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   3
      Left            =   2040
      Picture         =   "VIMMode2.frx":1C14
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   0
      Left            =   600
      Picture         =   "VIMMode2.frx":2056
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   2
      Left            =   1560
      Picture         =   "VIMMode2.frx":2498
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   1
      Left            =   1080
      Picture         =   "VIMMode2.frx":28DA
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmInterr2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DéjàFait As Boolean

Private Sub SScmdOK_Click()
Dim Début As String, Fautes As Integer, Eléments As String
    On Error Resume Next
    
    Faux = False
    
    ' Enlève
    For i% = 0 To 3
    For j% = 1 To 16
        Elément(i%, j%) = ""
    Next
    Next
    
    ' Sépares les éléments
    For e% = 0 To 3
        Eléments = UCase(Trim(Mid(Vrb(VrbEnCours), e% * 30 + 1, 30)))
        j% = 1
        For i% = 1 To Len(Eléments)
            If Mid(Eléments, i%, 1) <> "," Then
                Elément(e%, j%) = Elément(e%, j%) & Mid(Eléments, i%, 1)
            Else
                Elément(e%, j%) = Trim(Elément(e%, j%))
                j% = j% + 1
            End If
        Next
        Elément(e%, j%) = Trim(Elément(e%, j%))
    Next
    
    ' Cherche dans ces éléments
    For e% = 0 To 3
        'Enregistre la réponse
        strTexte$ = UCase(Trim(Texte(e%)))
        'Efface la réponse
        If Not Texte(e%).Locked Then Texte(e%).Text = ""
        ' Teste
        If strTexte$ = UCase(Trim(Mid(Vrb(VrbEnCours), e% * 30 + 1, 30))) Then GoTo ElémentSuite
        For i% = 1 To 16
            'Sort si vide
            If Elément(e%, i%) = "" Then Exit For
            If strTexte$ = Elément(e%, i%) Then GoTo ElémentSuite
        Next
        ' Faux
        If frmMode!optNoter = False Then
            Texte(e%).BackColor = QBColor(12)
            Texte(e%).Refresh
            Fautes = Fautes + 1
        Else
            ttFautes = ttFautes + 1
            Call NewVrb
            Exit Sub
        End If
ElémentSuite:
    Next

    If Fautes <> 0 And frmMode!optNoter = False Then
        Beep
        
        'Pour note finale
        ttFautes = ttFautes + 1
        
        'Redemande 1 Fois en plus
        If FoisInterr(VrbEnCours) < 3 And frmMode!optRedemande Then FoisInterr(VrbEnCours) = FoisInterr(VrbEnCours) + Int(Fautes / 2) + 1
        
        If Texte(0).Locked Then
            Texte(1).SetFocus
        Else
            Texte(0).SetFocus
        End If
    Else
        Call NewVrb
    End If
End Sub

Private Sub SScmdQuitter_Click()
    For i% = 1 To NbrVrb
        FoisInterr(i%) = 1
    Next
    
    DéjàFait = False
    Unload Me
End Sub

Private Sub cmdRéponce_Click()
    'Donne la réponce
    For i% = 0 To 3
        Texte(i%) = Trim(Mid(Vrb(VrbEnCours), i% * 30 + 1, 30))
    Next

    'Redemande 1 Fois en plus
    If FoisInterr(nbr) < 3 And frmMode!optRedemande Then FoisInterr(nbr) = FoisInterr(nbr) + 1
End Sub

Private Sub SScmdVFr_Click()
    ' Demande
    frmRecherche.Show 1
    
    'Efface les zones de textes
    For j% = 0 To 3
        Texte(j%).Text = ""
    Next

    'Créé un nbr aléatoire entre 0 et 3
    TexteAléatoire = Int(4 * Rnd)

    'Affiche le verbe
    Texte(TexteAléatoire).Text = Mid(Vrb(VrbEnCours), TexteAléatoire * 30 + 1, 30)
    Texte(TexteAléatoire).Locked = True

    'Donne le focus
    If Texte(0).Locked Then
        Texte(1).SetFocus
    Else
        Texte(0).SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    ' Test: Noter
    If frmMode!optNoter Then
        cmdRéponce.Enabled = False
        SScmdVFr.Enabled = False
        Me.Caption = "Interrogation Notée"
    End If
    
    ' Charge
    Tmp$ = ChargeListeVrb
    If Tmp$ = "cdlCancel" Then
        Hide
        Exit Sub
    Else
        Me.Icon = icoLangue(Val(Tmp$)).Picture
    End If

    ' Test: si Noter NbrVrb < 20
    If frmMode!optNoter And frmMode!chk20vrb And NbrVrb < 20 Then
        MsgBox "Cette liste de verbes ne contient que " & NbrVrb & " verbes." & Enter & _
            "Vous devez en choisir une autre pour l'interrogation sur 20 verbes!", vbInformation
        DéjàFait = False
    End If
    
    ' Affiche
    Call NewVrb
End Sub

Private Sub Noter()
Dim Note As Integer
    ' Donne la note
    If frmMode!optNoter And frmMode!chk20vrb Then NbrVrb = 20
    Note = 20 - ttFautes * 20 / NbrVrb
    
    ' Donne le message
    Select Case Note
        Case Is > 17
            Msg$ = "Très Bien: "
        Case Is > 14
            Msg$ = "Bien: "
        Case Is > 10
            Msg$ = "A Bien: "
        Case Is > 5
            Msg$ = "Il faut encore s'entrainer: "
        Case Else
            Msg$ = "NUL:"
    End Select
    
    ' Affiche
    ' Test: nbr entier
    If Int(Note) = Note Then
        Msg$ = Msg$ & Format(Note, "#0") & "/20"
    Else
        Msg$ = Msg$ & Format(Note, "#0.#") & "/20"
    End If
    MsgBox Msg$
End Sub

Public Sub NewVrb()
Dim TexteAléatoire As Integer
    On Error Resume Next
    
    ' Choisi un verbe aléatoirement dans la liste
    If -frmMode!chkAléatoire Then
        Tmp$ = VrbAléatoire
        If Tmp$ = "Recommencer" Then
            Form_Activate
        ElseIf Tmp$ = "Pas Recommencer" Then
            Hide
            Exit Sub
        End If
    Else
        VrbEnCours = VrbEnCours + 1
    End If
    
    ' Créé un nbr aléatoire entre 0 et 3
    TexteAléatoire = Int(4 * Rnd)

    ' Efface
    For i% = 0 To 3
        Texte(i%).Locked = False
        Texte(i%).Text = ""
    Next
    
    ' Affiche le verbe
    Texte(TexteAléatoire).Text = Trim(Mid(Vrb(VrbEnCours), TexteAléatoire * 30 + 1, 30))
    Texte(TexteAléatoire).Locked = True

    'Donne le focus
    If Texte(0).Locked Then
        Texte(1).SetFocus
    Else
        Texte(0).SetFocus
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SScmdQuitter_Click
End Sub

Private Sub Texte_Change(Index As Integer)
    ' Enlève la couleur de font
    Texte(Index).BackColor = QBColor(15)
End Sub

Private Sub Texte_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Sort si pas Echap
    If Not KeyAscii = 13 Then Exit Sub
    KeyAscii = 0
    
    'Donne le focus
    If Index = 3 Then
        SScmdOK_Click
    Else
        If Texte(Index + 1).Locked Then
            If Index + 2 >= 4 Then
                SScmdOK_Click
            Else
                Texte(Index + 2).SetFocus
            End If
        Else
            Texte(Index + 1).SetFocus
        End If
    End If
End Sub
