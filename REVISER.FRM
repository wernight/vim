VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmRéviser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Révision"
   ClientHeight    =   1890
   ClientLeft      =   600
   ClientTop       =   2835
   ClientWidth     =   9255
   Icon            =   "Reviser.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   375
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   8
      Top             =   360
      Width           =   1935
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
      Height          =   375
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   9
      Top             =   360
      Width           =   1935
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
      Height          =   375
      Index           =   2
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   30
      TabIndex        =   10
      Top             =   360
      Width           =   1935
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
      Height          =   375
      Index           =   3
      Left            =   6000
      Locked          =   -1  'True
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   3135
   End
   Begin Threed.SSCommand SScmdSuivant 
      Default         =   -1  'True
      Height          =   405
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
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
      Picture         =   "Reviser.frx":030A
   End
   Begin Threed.SSCommand SScmdPrécé 
      Height          =   405
      Left            =   480
      TabIndex        =   0
      Top             =   1200
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
      Enabled         =   0   'False
      BevelWidth      =   0
      Outline         =   0   'False
      AutoSize        =   2
      Picture         =   "Reviser.frx":0910
   End
   Begin Threed.SSCommand SScmdQuitter 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   7800
      TabIndex        =   3
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
      Picture         =   "Reviser.frx":0F1E
   End
   Begin Threed.SSCommand SScmdRecherche 
      Height          =   405
      Left            =   3720
      TabIndex        =   2
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
      Picture         =   "Reviser.frx":165E
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Infinitif"
      Height          =   195
      Index           =   0
      Left            =   -120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Traduction"
      Height          =   195
      Index           =   3
      Left            =   6000
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Participe passé"
      Height          =   195
      Index           =   2
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prétérit"
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image imgSuivant 
      Height          =   375
      Index           =   0
      Left            =   1920
      Picture         =   "Reviser.frx":21CA
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgSuivant 
      Height          =   375
      Index           =   1
      Left            =   1920
      Picture         =   "Reviser.frx":27C0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgPrécé 
      Height          =   375
      Index           =   0
      Left            =   480
      Picture         =   "Reviser.frx":2DC2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image imgPrécé 
      Height          =   375
      Index           =   1
      Left            =   480
      Picture         =   "Reviser.frx":33B0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   1
      Left            =   960
      Picture         =   "Reviser.frx":39AE
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   2
      Left            =   1440
      Picture         =   "Reviser.frx":3DF0
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "Reviser.frx":4232
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   3
      Left            =   1920
      Picture         =   "Reviser.frx":4674
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmRéviser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DéjàFait As Boolean, Nombre As Integer

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    ' Charge
    Tmp$ = ChargeListeVrb
    If Tmp$ = "cdlCancel" Then
        Hide
        Exit Sub
    Else
        Me.Icon = icoLangue(Val(Tmp$)).Picture
    End If

    ' Demande si aléatoire
    SScmdPrécé.Visible = Not -frmModeRévision!chkAléatoire
    SScmdSuivant_Click
End Sub
Private Sub SScmdPrécé_Click()
    ' Test si invisible
    If SScmdPrécé.Visible Then
        Nombre = Nombre - 2
        SScmdSuivant_Click
    End If
End Sub
Private Sub SScmdQuitter_Click()
    DéjàFait = False
    Unload Me
End Sub
Private Sub SScmdRecherche_Click()
    ' Demande
    frmRecherche.Show 1
    
    ' Affiche le verbe
    For i% = 0 To 3
        Texte(i%) = Trim(Mid(Vrb(VrbEnCours), i% * 30 + 1, 30))
    Next
    'Enlève Précé.
    SScmdPrécé.Enabled = VrbEnCours > 1
    SScmdPrécé.Picture = imgPrécé((VrbEnCours > 1) + 1).Picture
    ' Enlève Suiv.
    SScmdSuivant.Enabled = Nombre < NbrVrb
    SScmdSuivant.Picture = imgSuivant((Nombre < NbrVrb) + 1).Picture
    
    Reponse = ""
End Sub
Private Sub SScmdSuivant_Click()
    ' Test si invisible
    If SScmdPrécé.Visible Then
        Nombre = Nombre + 1
        'Affiche le verbe
        For j% = 0 To 3
            Texte(j%) = Trim(Mid(Vrb(Nombre), j% * 30 + 1, 30))
        Next
        'Enlève Précé.
        SScmdPrécé.Enabled = Nombre > 1
        SScmdPrécé.Picture = imgPrécé((Nombre > 1) + 1).Picture
    Else
        ' Choisi un verbe aléatoirement dans la liste
        Tmp$ = VrbAléatoire
        If Tmp$ = "Recommencer" Then
            Form_Activate
        ElseIf Tmp$ = "Pas Recommencer" Then
            Hide
            Exit Sub
        End If
    
        'Affiche le verbe
        For j% = 0 To 3
            Texte(j%) = Trim(Mid(Vrb(VrbEnCours), j% * 30 + 1, 30))
        Next
    End If
    
    ' Enlève Suiv.
    SScmdSuivant.Enabled = Nombre < NbrVrb
    SScmdSuivant.Picture = imgSuivant((Nombre < NbrVrb) + 1).Picture
End Sub

Private Sub Texte_Change(Index As Integer)
    Texte(Index).Refresh
End Sub
