VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Éditeur de liste de verbes irréguliers"
   ClientHeight    =   6285
   ClientLeft      =   750
   ClientTop       =   1905
   ClientWidth     =   9060
   Icon            =   "Edit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MouseIcon       =   "Edit.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   604
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanelListe 
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   645
      Width           =   9060
      _Version        =   65536
      _ExtentX        =   15981
      _ExtentY        =   1508
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   6
      Begin VB.TextBox Texte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         MaxLength       =   30
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Texte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   9
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Texte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3480
         MaxLength       =   30
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Texte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   11
         Top             =   360
         Width           =   3660
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prétérit"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   5
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Participe passé"
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   6
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Traduction"
         Height          =   255
         Index           =   4
         Left            =   5280
         TabIndex        =   7
         Top             =   120
         Width           =   3660
      End
      Begin VB.Label Label 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Infinitif"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4800
      Index           =   0
      Left            =   7335
      TabIndex        =   12
      Top             =   1470
      Width           =   1725
      Begin VB.Frame Frame 
         Caption         =   "Langue"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3225
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   165
         Width           =   1485
         Begin VB.OptionButton optLangue 
            Caption         =   "Allemand"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   21
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton optLangue 
            Caption         =   "Espagnol"
            Height          =   255
            Index           =   3
            Left            =   105
            TabIndex        =   20
            Top             =   1680
            Width           =   1215
         End
         Begin VB.OptionButton optLangue 
            Caption         =   "Italien"
            Height          =   255
            Index           =   2
            Left            =   105
            TabIndex        =   19
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton optLangue 
            Caption         =   "Anglais"
            Height          =   255
            Index           =   0
            Left            =   105
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   1215
         End
         Begin Threed.SSPanel SSPanelLangue 
            Height          =   735
            Left            =   375
            TabIndex        =   22
            Top             =   2280
            Width           =   735
            _Version        =   65536
            _ExtentX        =   1296
            _ExtentY        =   1296
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   2
            Begin VB.Image icoLangue 
               Height          =   480
               Index           =   10
               Left            =   120
               Picture         =   "Edit.frx":0614
               Top             =   120
               Width           =   480
            End
         End
         Begin VB.Image icoLangue 
            Height          =   480
            Index           =   0
            Left            =   -45
            Picture         =   "Edit.frx":091E
            Top             =   1995
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image icoLangue 
            Height          =   480
            Index           =   1
            Left            =   -45
            Picture         =   "Edit.frx":0D60
            Top             =   2475
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image icoLangue 
            Height          =   480
            Index           =   2
            Left            =   1110
            Picture         =   "Edit.frx":11A2
            Top             =   2055
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image icoLangue 
            Height          =   480
            Index           =   3
            Left            =   1110
            Picture         =   "Edit.frx":15E4
            Top             =   2505
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame fraCharacSpéc 
         Caption         =   "Caractères spéciaux"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         TabIndex        =   15
         Top             =   3450
         Width           =   1485
         Begin VB.Label lblCharacSpéc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Aucun"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   465
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Nombre de verbes"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   13
         Top             =   4170
         Width           =   1485
         Begin VB.Label lblNbrVrb 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            Height          =   195
            Left            =   300
            TabIndex        =   14
            Top             =   180
            Width           =   630
         End
      End
   End
   Begin VB.ListBox lstVrb 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1020
      IntegralHeight  =   0   'False
      ItemData        =   "Edit.frx":1A26
      Left            =   120
      List            =   "Edit.frx":1A28
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.ListBox Liste 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      IntegralHeight  =   0   'False
      ItemData        =   "Edit.frx":1A2A
      Left            =   0
      List            =   "Edit.frx":1A2C
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1545
      Width           =   7305
   End
   Begin Threed.SSPanel SSPanelListe 
      Height          =   1485
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9060
      _Version        =   65536
      _ExtentX        =   15981
      _ExtentY        =   2619
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   0
      BevelOuter      =   1
      BevelInner      =   2
      Alignment       =   6
      Begin VB.Image imgOuvrirDn 
         Height          =   600
         Left            =   1155
         Picture         =   "Edit.frx":1A2E
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdOuvrir 
         Height          =   600
         Left            =   1155
         Picture         =   "Edit.frx":1D20
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgOuvrirUp 
         Height          =   600
         Left            =   1155
         Picture         =   "Edit.frx":1FBC
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgQuitterDn 
         Height          =   600
         Left            =   6780
         Picture         =   "Edit.frx":22A4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgQuitterUp 
         Height          =   600
         Left            =   6780
         Picture         =   "Edit.frx":26AC
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgQuitter 
         Height          =   600
         Left            =   6780
         Picture         =   "Edit.frx":2A9E
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdQuitter 
         Height          =   600
         Left            =   6780
         Picture         =   "Edit.frx":2E50
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgEnregDn 
         Height          =   600
         Left            =   2280
         Picture         =   "Edit.frx":3202
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgEnregUp 
         Height          =   600
         Left            =   2280
         Picture         =   "Edit.frx":358A
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdEnreg 
         Height          =   600
         Left            =   2280
         Picture         =   "Edit.frx":38F4
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgEnreg 
         Height          =   600
         Left            =   2280
         Picture         =   "Edit.frx":3C12
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgNewDn 
         Height          =   600
         Left            =   30
         Picture         =   "Edit.frx":3F30
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgNewUp 
         Height          =   600
         Left            =   30
         Picture         =   "Edit.frx":423A
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgNew 
         Height          =   600
         Left            =   30
         Picture         =   "Edit.frx":4540
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdNew 
         Height          =   600
         Left            =   30
         Picture         =   "Edit.frx":47FA
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgInserDn 
         Height          =   600
         Left            =   3405
         Picture         =   "Edit.frx":4AB4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgInserUp 
         Height          =   600
         Left            =   3405
         Picture         =   "Edit.frx":4E12
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgSupprDn 
         Height          =   600
         Left            =   4530
         Picture         =   "Edit.frx":5160
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgSupprUp 
         Height          =   600
         Left            =   4530
         Picture         =   "Edit.frx":58F6
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgInser 
         Height          =   600
         Left            =   3405
         Picture         =   "Edit.frx":608C
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdInser 
         Height          =   600
         Left            =   3405
         Picture         =   "Edit.frx":673A
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgAideDn 
         Height          =   600
         Left            =   7905
         Picture         =   "Edit.frx":6DE8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgAideUp 
         Height          =   600
         Left            =   7905
         Picture         =   "Edit.frx":74D8
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgAide 
         Height          =   600
         Left            =   7905
         Picture         =   "Edit.frx":7BC8
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdAide 
         Height          =   600
         Left            =   7905
         Picture         =   "Edit.frx":821E
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image cmdSuppr 
         Height          =   600
         Left            =   4530
         Picture         =   "Edit.frx":8874
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgSuppr 
         Height          =   600
         Left            =   4530
         Picture         =   "Edit.frx":8F70
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image cmdAnnuler 
         Height          =   600
         Left            =   5655
         Picture         =   "Edit.frx":966C
         Top             =   30
         Width           =   1125
      End
      Begin VB.Image imgAnnulerDn 
         Height          =   600
         Left            =   5655
         Picture         =   "Edit.frx":9D1C
         Top             =   1080
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgAnnulerUp 
         Height          =   600
         Left            =   5655
         Picture         =   "Edit.frx":A464
         Top             =   840
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgAnnuler 
         Height          =   600
         Left            =   5655
         Picture         =   "Edit.frx":ABAC
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Image imgOuvrir 
         Height          =   600
         Left            =   1155
         Picture         =   "Edit.frx":B25C
         Top             =   600
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Menu menuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu mnuNouveau 
         Caption         =   "&Nouveau"
      End
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Ouvrir"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSéparateur1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnreg 
         Caption         =   "&Enregistrer"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuEnregSous 
         Caption         =   "En&registrer sous..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSéparateur2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintCfg 
         Caption         =   "&Configuration de l'impression..."
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Imprimer..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSéparateur3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu menuEdition 
      Caption         =   "&Edition"
      Begin VB.Menu mnuAnnuler 
         Caption         =   "Impossible d'annuler"
         Enabled         =   0   'False
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSéparateur4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInser 
         Caption         =   "&Insérer"
      End
      Begin VB.Menu mnuSuppr 
         Caption         =   "&Supprimer"
      End
   End
   Begin VB.Menu menuAide 
      Caption         =   "&?"
      Begin VB.Menu mnuApropos 
         Caption         =   "A &propos"
      End
      Begin VB.Menu mnuCommander 
         Caption         =   "&Commander"
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Langue As Integer, LignesSuppr(0 To 500) As String
Dim TxtChange As Boolean, ListeChange As Boolean

Private Sub cmdAide_Click()
    cmdAide.Refresh

    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        MsgBox "Cliquer sur un des boutons pour" & Enter & _
        "avoir une aide dessus.", , "Aide sur 'AIDE'"
        Exit Sub
    End If
    
    #If Win16 Then
        Me.MousePointer = 99
    #Else
        Me.MousePointer = 14
    #End If
End Sub

Private Sub cmdAide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdAide = imgAideDn
End Sub

Private Sub cmdAide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdAide.Width Or Y < 0 Or Y > cmdAide.Height) Then
            cmdAide = imgAideUp
        End If
    Case 1
        If X <= 0 Or X > cmdAide.Width Or Y < 0 Or Y > cmdAide.Height Then
            cmdAide = imgAide
        Else
            cmdAide = imgAideDn
        End If
    End Select
End Sub

Private Sub cmdAide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdAide = imgAide
End Sub

Private Sub cmdAnnuler_Click()
    cmdAnnuler.Refresh
    
    If mnuAnnuler.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    mnuAnnuler_Click
    
    ' Affiche le nbr de verbes
    AffNbrVrb
End Sub

Private Sub cmdAnnuler_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuAnnuler.Enabled = False And Me.MousePointer = 0 Then
        cmdAnnuler = imgAnnuler
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdAnnuler.Width Or Y < 0 Or Y > cmdAnnuler.Height) Then
            cmdAnnuler = imgAnnulerUp
        End If
    Case 1
        If X <= 0 Or X > cmdAnnuler.Width Or Y < 0 Or Y > cmdAnnuler.Height Then
            cmdAnnuler = imgAnnuler
        Else
            cmdAnnuler = imgAnnulerDn
        End If
    End Select
End Sub

Private Sub cmdAnnuler_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdAnnuler = imgAnnuler
End Sub

Private Sub cmdOuvrir_Click()
    cmdOuvrir.Refresh
    
    If cmdOuvrir.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    mnuOuvrir_Click
End Sub

Private Sub cmdOuvrir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuOuvrir.Enabled = False Then
        cmdOuvrir = imgOuvrir
        Exit Sub
    End If
    
    If Button = 1 Then cmdOuvrir = imgOuvrirDn
End Sub

Private Sub cmdOuvrir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuOuvrir.Enabled = False And Me.MousePointer = 0 Then
        cmdOuvrir = imgOuvrir
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdOuvrir.Width Or Y < 0 Or Y > cmdOuvrir.Height) Then
            cmdOuvrir = imgOuvrirUp
        End If
    Case 1
        If X <= 0 Or X > cmdOuvrir.Width Or Y < 0 Or Y > cmdOuvrir.Height Then
            cmdOuvrir = imgOuvrir
        Else
            cmdOuvrir = imgOuvrirDn
        End If
    End Select
End Sub

Private Sub cmdOuvrir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdOuvrir = imgOuvrir
End Sub

Private Sub fraCharacSpéc_Click()
Dim Msg As String
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Dans certaines langues il existe des charactères spéciaux." & Enter & _
        "Pour les utiliser, appuyez sur ALT et entrez le code à" & Enter & _
        "l 'aide des touches numériques de droite.", , "Aide sur 'Craractères Spéciaux'"
        Exit Sub
    End If
End Sub

Private Sub cmdEnreg_Click()
    cmdEnreg.Refresh
    
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Enregistre la liste.", , "Aide sur 'Enregistrer'"
        Exit Sub
    End If
    
    If cmdEnreg.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    
    mnuEnreg_Click
End Sub

Private Sub cmdEnreg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuEnreg.Enabled = False Then
        cmdEnreg = imgEnreg
        Exit Sub
    End If
    
    If Button = 1 Then cmdEnreg = imgEnregDn
End Sub

Private Sub cmdEnreg_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuEnreg.Enabled = False And Me.MousePointer = 0 Then
        cmdEnreg = imgEnreg
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdEnreg.Width Or Y < 0 Or Y > cmdEnreg.Height) Then
            cmdEnreg = imgEnregUp
        End If
    Case 1
        If X <= 0 Or X > cmdEnreg.Width Or Y < 0 Or Y > cmdEnreg.Height Then
            cmdEnreg = imgEnreg
        Else
            cmdEnreg = imgEnregDn
        End If
    End Select
End Sub

Private Sub cmdEnreg_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdEnreg = imgEnreg
End Sub

Private Sub cmdInser_Click()
    cmdInser.Refresh
    
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Ajoute une nouvelle ligne dans la liste.", , "Aide sur 'Insérer'"
        Exit Sub
    End If
    
    If cmdInser.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    
    ' Test si texte à changé
    If TxtChange Then Call Enreg(Liste.ListIndex)
    
    ' Enlève la propriétée Locked
    For i% = 1 To 4
        Texte(i%).Locked = False
    Next
    
    ' Ajoute une nouvelle ligne
    Liste.AddItem "", Liste.ListIndex
    lstVrb.AddItem "", Liste.ListIndex
    
    ' Sélectionne cette ligne
    If Liste.ListCount <> 1 Then
        Liste.Selected(Liste.ListIndex + 1) = False
    End If
    Liste.Selected(Liste.ListIndex) = True

    ' Efface le texte
    For i% = 1 To 4
        Texte(i%) = ""
    Next
    
    ' Donne le focus
    Texte(1).SetFocus
    
    ' Affiche le nbr de verbes
    AffNbrVrb
End Sub

Private Sub cmdinser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuInser.Enabled = False Then
        cmdInser = imgInser
        Exit Sub
    End If
    
    If Button = 1 Then cmdInser = imgInserDn
End Sub

Private Sub cmdinser_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuInser.Enabled = False And Me.MousePointer = 0 Then
        cmdInser = imgInser
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdInser.Width Or Y < 0 Or Y > cmdInser.Height) Then
            cmdInser = imgInserUp
        End If
    Case 1
        If X <= 0 Or X > cmdInser.Width Or Y < 0 Or Y > cmdInser.Height Then
            cmdInser = imgInser
        Else
            cmdInser = imgInserDn
        End If
    End Select
End Sub

Private Sub cmdinser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdInser = imgInser
End Sub

Private Sub cmdNew_Click()
    cmdNew.Refresh
    
    If mnuNouveau.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    mnuNouveau_Click
End Sub

Private Sub cmdNew_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuNouveau.Enabled = False Then
        cmdNew = imgNew
        Exit Sub
    End If
    
    If Button = 1 Then cmdNew = imgNewDn
End Sub

Private Sub cmdNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuNouveau.Enabled = False And Me.MousePointer = 0 Then
        cmdNew = imgNew
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdNew.Width Or Y < 0 Or Y > cmdNew.Height) Then
            cmdNew = imgNewUp
        End If
    Case 1
        If X <= 0 Or X > cmdNew.Width Or Y < 0 Or Y > cmdNew.Height Then
            cmdNew = imgNew
        Else
            cmdNew = imgNewDn
        End If
    End Select
End Sub

Private Sub cmdNew_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdNew = imgNew
End Sub

Private Sub cmdQuitter_Click()
    cmdQuitter.Refresh
    
    If mnuQuitter.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    mnuQuitter_Click
End Sub

Private Sub cmdQuitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuQuitter.Enabled = False Then
        cmdQuitter = imgQuitter
        Exit Sub
    End If
    
    If Button = 1 Then cmdQuitter = imgQuitterDn
End Sub

Private Sub cmdQuitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuQuitter.Enabled = False And Me.MousePointer = 0 Then
        cmdQuitter = imgQuitter
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdQuitter.Width Or Y < 0 Or Y > cmdQuitter.Height) Then
            cmdQuitter = imgQuitterUp
        End If
    Case 1
        If X <= 0 Or X > cmdQuitter.Width Or Y < 0 Or Y > cmdQuitter.Height Then
            cmdQuitter = imgQuitter
        Else
            cmdQuitter = imgQuitterDn
        End If
    End Select
End Sub

Private Sub cmdQuitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdQuitter = imgQuitter
End Sub

Private Sub cmdSuppr_Click()
    cmdSuppr.Refresh
    
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Supprime les lignes sélectionnées.", , "Aide sur 'Supprimer'"
        Exit Sub
    End If
    
    If cmdSuppr.Enabled = False And Me.MousePointer = 0 Then Exit Sub
    
    ' Test si texte à changé
    If TxtChange Then Call Enreg(Liste.ListIndex)
    ListeChange = True
    
    ' Test si dernière ligne
    If Liste.ListIndex = Liste.ListCount - 1 Then Exit Sub
    
    ' Efface le texte
    For i% = 1 To 4
        Texte(i%) = ""
    Next
    
    ' Affiche pour Annuler
    mnuAnnuler.Caption = "&Annuler Effacer"
    mnuAnnuler.Enabled = True
    
    ' Enlève l'ancien Annuler
    s% = 0
    Do Until LignesSuppr(s%) = ""
        LignesSuppr(s%) = ""
        s% = s% + 1
    Loop
    
    ' Test, Enregistre, Efface...
    i% = 0
    e% = 0
    Do While i% < Liste.ListCount
        If Liste.Selected(i%) Then
            ' Enregistre pour Annuler
            LignesSuppr(s%) = i% + s% & lstVrb.List(i%)
            s% = s% + 1
            ' Efface
            Liste.RemoveItem i%
            lstVrb.RemoveItem i%
        Else
            i% = i% + 1         'Change de ligne
        End If
    Loop
    
    ' Donne la possibilitée de sélectionner
    Liste.Enabled = True
    
    ' Enlève texte change
    TxtChange = False
    
    ' Affiche le nbr de verbes
    AffNbrVrb
End Sub

Private Sub cmdSuppr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuSuppr.Enabled = False Then
        cmdSuppr = imgSuppr
        Exit Sub
    End If
    
    If Button = 1 Then cmdSuppr = imgSupprDn
End Sub

Private Sub cmdSuppr_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
    
    If mnuSuppr.Enabled = False And Me.MousePointer = 0 Then
        cmdSuppr = imgSuppr
        Exit Sub
    End If
    
    Select Case Button
    Case 0
        If Not (X <= 0 Or X > cmdSuppr.Width Or Y < 0 Or Y > cmdSuppr.Height) Then
            cmdSuppr = imgSupprUp
        End If
    Case 1
        If X <= 0 Or X > cmdSuppr.Width Or Y < 0 Or Y > cmdSuppr.Height Then
            cmdSuppr = imgSuppr
        Else
            cmdSuppr = imgSupprDn
        End If
    End Select
End Sub

Private Sub cmdSuppr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then cmdSuppr = imgSuppr
End Sub

Private Sub Form_Load()
Dim Fichier As String, Charactère As String
    On Error Resume Next
    
    TxtChange = False
    ListeChange = False
    
    ' Change la config
    frmVIM!CommonDialog.DialogTitle = "Modifier une liste de verbe"

    ' Affiche
    Call AffListe
End Sub

Private Sub AffListe()
Dim Verbe As String, Mot2 As String * 15
    On Error GoTo ErrHandler
    
    ' Efface
    For i% = 1 To 4
        Texte(i%) = ""
    Next
    Liste.Clear
    lstVrb.Clear
    
    ListeChange = False
    TxtChange = False
    
    ' Demande le fichier
    frmVIM!CommonDialog.FileName = ListVrb
    frmVIM!CommonDialog.ShowOpen
    ListVrb = frmVIM!CommonDialog.FileName
    
    ' Charge
    Tmp$ = ChargeListeVrb
    Langue = Val(Tmp$)
    optLangue(Langue) = True
    Me.Icon = icoLangue(Langue)
    icoLangue(10) = icoLangue(Langue)
    
    ' Affiche
    For i% = 1 To NbrVrb
        lstVrb.AddItem Vrb(i%)
        ' N'affiche que les 15-1er charac.
        Verbe = ""
        For j% = 0 To 2
            Mot2 = Mid(Vrb(i%), j% * 30 + 1, 15)
            Verbe = Verbe & Mot2 & " "
        Next
        Verbe = Verbe & Mid(Vrb(i%), j% * 30 + 1, 30)
        Liste.AddItem Verbe
    Next
    
    ' Ajoute une ligne vide
    Liste.AddItem "", Liste.ListCount
    lstVrb.AddItem "", lstVrb.ListCount
    If Liste.ListCount = 1 Then Liste.AddItem ""
    
    ' Sélectionne
    Liste.Selected(0) = True
    
    ' Enlève Enabled
    Call FoncEnabled(Liste.List(1) <> "")
    
    GoTo Fin
    
ErrHandler:
    Close
    If Err = 340 Then
        MsgBox "Liste de verbes non valide.", vbExclamation
    ElseIf Err = cdlCancel Then
        Unload frmEdit
        Exit Sub
    Else
        MsgBox "L'erreur suivante c'est produite:" & Enter & Error(Err)
    End If
    Resume Fin
    
Fin:
    ' Donne le nom du fichier
    Fichier = ""
    For i% = 1 To 12
        Charactère = Left(Right(frmVIM!CommonDialog.FileName, i%), 1)
        If Charactère = "\" Then Exit For
        Fichier = Charactère & Fichier
    Next
    Me.Caption = Fichier & " - Liste de verbes"
    
    ListeChange = False
    TxtChange = False
    
    
    ' Affiche le nbr de verbes
    AffNbrVrb
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNew = imgNew
    cmdOuvrir = imgOuvrir
    cmdEnreg = imgEnreg
    cmdInser = imgInser
    cmdSuppr = imgSuppr
    cmdAnnuler = imgAnnuler
    cmdQuitter = imgQuitter
    cmdAide = imgAide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmVIM!CommonDialog.DialogTitle = "Ouvrir une liste de verbes"
End Sub

Private Sub icoLangue_Click(Index As Integer)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Affiche l'icone de la langue en cours.", , "Aide sur 'Langue'"
        Exit Sub
    End If
End Sub

Private Sub Liste_Click()
Static LigneAv As Integer
Dim Mot As String, TxtLock As Boolean
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Liste de verbes." & Enter & "Utilisez les zones de textes au dessus" & Enter & _
        "de celle-ci pour modifier la liste.", , "Aide sur 'Liste'"
        Exit Sub
    End If
    
    ' Test si texte à changé
    If TxtChange And Liste.ListCount > 1 Then Call Enreg(LigneAv)
    
    ' Test si dernière ligne
    TxtLock = (Liste.ListIndex = Liste.ListCount - 1)
   
    ' Locked
    cmdSuppr.Enabled = Not TxtLock
    For i% = 1 To 4
        Texte(i%).Locked = TxtLock
    Next
    
    ' Sépare les mots de la ligne
    For i% = 0 To 3
        Mot = Trim(Mid(lstVrb.List(Liste.ListIndex), i% * 30 + 1, 30))
        Texte(i% + 1) = Mot
    Next
    
    ' Enregistre
    LigneAv = Liste.ListIndex
    TxtChange = False
End Sub

Private Sub Liste_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Teste Suppr
    If KeyCode = 46 And Shift = 0 Then cmdSuppr_Click
End Sub

Private Sub Liste_KeyPress(KeyAscii As Integer)
    ' Test
    If Texte(1).Locked Then
        cmdInser_Click
        Texte(1) = Texte(1) & Chr(KeyAscii)
        Texte(1).SelStart = 1
        KeyAscii = 0
    ElseIf Liste.ListIndex > Liste.ListCount - 2 Then
        Exit Sub
    ElseIf KeyAscii = 13 Then
        Liste.Selected(Liste.ListIndex) = False
        Liste.Selected(Liste.ListIndex + 1) = True
    End If
End Sub

Private Sub mnuAnnuler_Click()
Dim Verbe As String, Texte As String, Mot As String * 15
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Annuler si vous avez supprimé des éléments dans la liste.", , "Aide sur 'Annuler'"
        Exit Sub
    End If
    
    ' Test si texte à changé
    If TxtChange Then Call Enreg(Liste.ListIndex)
    
    ' Annule
    s% = 0
    Do Until LignesSuppr(s%) = ""
        Texte = Right(LignesSuppr(s%), 120)
        'N'affiche que 15-1er Charac
        Verbe = ""
        For i% = 0 To 3
            Mot2 = Mid(Texte, i% * 30 + 1, 15)
            Verbe = Verbe & Mot2 & " "
        Next
        Verbe = Verbe & Mid(Texte, i% * 30 + 1, 30)
        ' Affiche
        Liste.AddItem Verbe, Left(LignesSuppr(s%), 1)
        lstVrb.AddItem Texte, Left(LignesSuppr(s%), 1)
        s% = s% + 1
    Loop
    
    ' Affiche
    mnuAnnuler.Caption = "Impossible d'annuler"
    mnuAnnuler.Enabled = False
End Sub

Private Sub mnuApropos_Click()
    frmAPropos.Show 1
End Sub

Private Sub mnuCommander_Click()
    frmReg.Show 1
End Sub

Private Sub mnuEnreg_Click()
    ' Enregistre
    ListeChange = False
    
    ' Test si texte à changé
    If TxtChange Then Call Enreg(Liste.ListIndex)
    
    ' Efface l'ancien fichier
    Kill frmVIM!CommonDialog.FileName
    
    ' Crée le fichier pour y rentrer les infos
    Open frmVIM!CommonDialog.FileName For Random As 1 Len = 122
    Put #1, , Langue
    For i% = 0 To Liste.ListCount - 2
        Put #1, , lstVrb.List(i%)
    Next
    Close
End Sub

Private Sub mnuEnregSous_Click()
    On Error Resume Next
    
    frmVIM!CommonDialog.DialogTitle = "Enregistrer sous"

    ' Enregistre
    ListeChange = False
    
    ' Test si texte à changé
    If TxtChange Then Call Enreg(Liste.ListIndex)
    
Demande:
    ' Demande le nom
    frmVIM!CommonDialog.ShowSave
    If Err = cdlCancel Then Exit Sub
    
    ' Efface l'ancien fichier
    If Dir(frmVIM!CommonDialog.FileName) <> "" Then
        Msg$ = frmVIM!CommonDialog.FileName & Enter & "Ce fichier existe déjà." & Enter & Enter & "Voulez-vous le remplacer?"
        If MsgBox(Msg$, vbExclamation + vbYesNo, "Enregistrer sous") = vbYes Then
            Kill frmVIM!CommonDialog.FileName    'Efface
        Else
            GoTo Demande            'Revient
        End If
    End If
    
    ' Crée le fichier pour y rentrer les infos
    Open frmVIM!CommonDialog.FileName For Random As #1 Len = 122
    Put #1, , Langue
    For i% = 0 To Liste.ListCount - 2
        Put #1, , lstVrb.List(i%)
    Next
    Close

    ' Donne le nom du fichier
    Fichier = ""
    For i% = 1 To 12
        Charactère = Left(Right(frmVIM!CommonDialog.FileName, i%), 1)
        If Charactère = "\" Then Exit For
        Fichier = Charactère & Fichier
    Next
    Me.Caption = Fichier & " - Liste de verbes"
End Sub

Private Sub mnuPrint_Click()
Rem Imprime la liste
    Load frmPrint
    Unload frmPrint
End Sub

Private Sub mnuInser_Click()
    cmdInser_Click
End Sub

Private Sub mnuNouveau_Click()
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Créé une nouvelle liste de verbes.", , "Aide sur 'Nouveau'"
        Exit Sub
    End If
    
    ' Test si enregistré
    If ListeChange Then
        ' Demande
        Réponse = MsgBox("Enregistrer les modifications sous '" & frmVIM!CommonDialog.FileName & "' ?", vbExclamation + vbYesNoCancel)
        Select Case Réponse
            Case vbYes
                mnuEnreg_Click
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    ' Mise à 0
    TxtChange = False
    ListeChange = False
    
    ' Efface
    Liste.Clear
    lstVrb.Clear
    For i% = 1 To 4
        Texte(i%) = ""
    Next
    
    ' Ajoute une ligne vide
    Liste.AddItem "", Liste.ListCount
    lstVrb.AddItem "", Liste.ListCount
    If Liste.ListCount = 1 Then Liste.AddItem ""
    
    ' Sélectionne
    Liste.Selected(0) = True
End Sub

Private Sub mnuOuvrir_Click()
    On Error Resume Next
    
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Ouvre une liste déjà existante pour en modifier le contenu.", , "Aide sur 'Ouvrir'"
        Exit Sub
    End If
    
    ' Test si enregistré
    If ListeChange Then
        ' Demande
        Réponse = MsgBox("Enregistrer les modifications sous '" & frmVIM!CommonDialog.FileName & "' ?", vbExclamation + vbYesNoCancel)
        Select Case Réponse
            Case vbYes
                mnuEnreg_Click
            Case vbCancel
                Exit Sub
        End Select
    End If
    
    ' Change la config
    frmVIM!CommonDialog.DialogTitle = "Modifier une liste de verbe"
    
    ' Affiche la liste de verbes
    Call AffListe
End Sub

Private Sub mnuPrintCfg_Click()
    On Error Resume Next
    
    frmVIM!CommonDialog.Flags = cdlPDPrintSetup
    frmVIM!CommonDialog.ShowPrinter
End Sub

Private Sub mnuQuitter_Click()
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Sort de la fenêtre pour revenir" & Enter & _
        "au menu principal.", , "Aide sur 'Quitter'"
        Exit Sub
    End If
    
    ' Test si enregistré
    If ListeChange Then
        ' Demande
        Réponse = MsgBox("Enregistrer les modifications sous '" & frmVIM!CommonDialog.FileName & "' ?", vbExclamation + vbYesNoCancel)
        Select Case Réponse
            Case vbYes
                mnuEnreg_Click
            Case vbCancel
                Exit Sub
        End Select
    End If

    ' Sort
    Hide
    frmVIM.Show
    Unload Me
End Sub

Private Sub mnuSuppr_Click()
    cmdSuppr_Click
End Sub

Private Sub optLangue_Click(Index As Integer)
    Langue = Index
    ' Donne les charactères Spéciaux
    If Langue = 1 Then
        lblCharacSpéc = "ß = 225 ou 0223"
    Else
        lblCharacSpéc = "Aucun"
    End If
    Me.Icon = icoLangue(Index).Picture
    icoLangue(10).Picture = icoLangue(Index).Picture
    ListeChange = True
End Sub

Private Sub cmdAnnuler_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mnuAnnuler.Enabled = False Then
        cmdAnnuler = imgAnnuler
        Exit Sub
    End If
    
    If Button = 1 Then cmdAnnuler = imgAnnulerDn
End Sub

Private Sub optLangue_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Choisissez la langue de la liste en cours. La langue est est destinée" & Enter & _
          "à vous informer sur la langue de la liste quand vous révisez ou" & Enter & _
          "êtes questionné.", , "Aide sur 'Langue'"
        Exit Sub
    End If
End Sub

Private Sub SSPanelListe_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(0, 0, 10, 10)
End Sub

Private Sub Texte_Change(Index As Integer)
    ' Test si dernière ligne
    If Texte(1).Locked = False Then
        ' Test ligne vide
        FoncEnabled (Texte(1) <> "" And Texte(4) <> "")
    End If
    
    ' Enregistre qu'il y a changement
    TxtChange = True
    ListeChange = True
End Sub

Private Sub FoncEnabled(Val As Boolean)
    ' Affiche les changements
    Liste.Enabled = Val
    cmdInser.Enabled = Val
    cmdEnreg.Enabled = Val
    mnuEnreg.Enabled = Val
    mnuEnregSous.Enabled = Val
End Sub

Private Sub Enreg(LigneAv As Integer)
Dim Verbe As String, Mot As String * 30, Mot2 As String * 15
    'Met en place le verbes
    For i% = 1 To 4
        Mot = Texte(i%)
        Verbe = Verbe & Mot
    Next
    
    'Affiche
    lstVrb.List(LigneAv) = Verbe
    
    'Met en place le verbes
    Verbe = ""
    For i% = 1 To 3
        Mot2 = Texte(i%)
        Verbe = Verbe & Mot2 & " "
    Next
    Verbe = Verbe & Texte(i%)
    
    'Affiche
    Liste.List(LigneAv) = Verbe
    Liste.Refresh
    
    TxtChange = False
End Sub

Private Sub Texte_Click(Index As Integer)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Vous permet de modifier l'élément sélectionné dans la liste. Vous povez donner plusieurs possibilités en les séparent par des virgules.", , "Aide sur 'Texte Modif'"
        Exit Sub
    End If
End Sub

Private Sub Texte_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Sort si dernière ligne
    If Liste.ListIndex > Liste.ListCount - 2 Then Exit Sub

    ' Teste pour MAJ
    If KeyAscii = 13 Then
        KeyAscii = 0
        ' Test si texte valide
        If Liste.Enabled = False Then
            If Index <> 4 Then Texte(Index + 1).SetFocus
            Exit Sub
        End If
        
        ' Enregistre
        Call Enreg(Liste.ListIndex)
        
        ' Test si le verbe existe déjà (double)
        Element$ = Trim(lstVrb.List(Liste.ListIndex))
        IndexListe% = Liste.ListIndex
        For i% = 0 To lstVrb.ListCount - 2
            If Trim(lstVrb.List(i%)) = Element$ Then
            If IndexListe% <> i% Then
                MsgBox "Ce verbe existe déjà!", vbExclamation
            End If
            End If
        Next
        
        ' Change de ligne
        Liste.Selected(Liste.ListIndex) = False
        Liste.Selected(Liste.ListIndex + 1) = True
    End If
End Sub

Private Sub tmrBulleAide_Timer()
    lblBulleAide.Visible = False
    tmrBulleAide.Enabled = False
End Sub

Public Sub AffNbrVrb()
    ' Affiche le nbr de verbes
    lblNbrVrb = Liste.ListCount - 1
End Sub
