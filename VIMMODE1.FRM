VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmInterr1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2130
   ClientLeft      =   1005
   ClientTop       =   3765
   ClientWidth     =   6975
   Icon            =   "VIMMode1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Texte 
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
      Left            =   4680
      MaxLength       =   30
      TabIndex        =   2
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      MaxLength       =   30
      TabIndex        =   0
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
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2520
      MaxLength       =   30
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
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
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin Threed.SSCommand SScmdVFr 
      Height          =   390
      Left            =   345
      TabIndex        =   9
      Top             =   1545
      Width           =   6285
      _Version        =   65536
      _ExtentX        =   11086
      _ExtentY        =   688
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
      Picture         =   "VIMMode1.frx":030A
   End
   Begin Threed.SSCommand SScmdQuitter 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   4665
      TabIndex        =   8
      Top             =   1080
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
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
      Picture         =   "VIMMode1.frx":10A0
   End
   Begin Threed.SSCommand SScmdOK 
      Height          =   405
      Left            =   2505
      TabIndex        =   7
      Top             =   1065
      Width           =   1965
      _Version        =   65536
      _ExtentX        =   3466
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
      Picture         =   "VIMMode1.frx":17E0
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   3
      Left            =   6600
      Picture         =   "VIMMode1.frx":1E3E
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   0
      Left            =   6600
      Picture         =   "VIMMode1.frx":2280
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   2
      Left            =   6600
      Picture         =   "VIMMode1.frx":26C2
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoLangue 
      Height          =   480
      Index           =   1
      Left            =   6600
      Picture         =   "VIMMode1.frx":2B04
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Infinitif"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Participe passé"
      Height          =   195
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prétérit"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmInterr1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DéjàFait As Boolean

Private Sub SScmdOK_Click()
Dim Début As String, Fautes As Integer
    On Error Resume Next
    
    Faux = False
    
    ' Enlève
    For i% = 0 To 3
    For j% = 1 To 16
        Elément(i%, j%) = ""
    Next
    Next
    
    ' Sépares les éléments
    For e% = 0 To 2
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
    For e% = 0 To 2
        'Enregistre la réponse
        strTexte$ = UCase(Trim(Texte(e%)))
        'Efface la réponse
        Texte(e%).Text = ""
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
        
        'Redemande 1 Fois en plus
        If FoisInterr(VrbEnCours) < 3 And frmMode!optRedemande Then FoisInterr(VrbEnCours) = FoisInterr(VrbEnCours) + Int(Fautes / 2) + 1
        
        Texte(1).SetFocus
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
    For i% = 0 To 2
        Texte(i%).Text = Trim(Mid(Vrb(VrbEnCours), i% * 30 + 1, 30))
    Next

    'Redemande 1 Fois en plus
    If FoisInterr(VrbEnCours) < 3 And frmMode!optRedemande Then FoisInterr(VrbEnCours) = FoisInterr(VrbEnCours) + 1
End Sub

Private Sub SScmdVFr_Click()
    ' Demande
    frmRecherche.Show 1
    
    'Affiche le verbe
    Me.Caption = "Verbe: " & Mid(Vrb(VrbEnCours), 90, 30)
    For j% = 0 To 2
        Texte(j%).Text = ""
    Next
    Texte(0).SetFocus
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    ' Test: Noter
    If frmMode!optNoter Then
        cmdRéponce.Enabled = False
        SScmdVFr.Enabled = False
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
        Texte(Index + 1).SetFocus
    End If
End Sub

Public Sub NewVrb()
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
    
    ' Affiche le verbe
    Me.Caption = Trim("Verbe: " & Mid(Vrb(VrbEnCours), 90, 30))
    For j% = 0 To 2
        Texte(j%).Text = ""
    Next
    Texte(0).SetFocus
End Sub
