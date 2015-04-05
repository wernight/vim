VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmMode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mode d'interrogation"
   ClientHeight    =   2550
   ClientLeft      =   2640
   ClientTop       =   2790
   ClientWidth     =   5415
   ControlBox      =   0   'False
   Icon            =   "Mode.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Mode.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2550
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFichier 
      Caption         =   "&Fichier"
      Height          =   285
      Left            =   4320
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtFichier 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Width           =   4080
   End
   Begin VB.CheckBox chkAléatoire 
      Caption         =   "Aléatoirement"
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame fraAutres 
      Caption         =   "Autres options :"
      Height          =   1335
      Left            =   2685
      TabIndex        =   7
      Top             =   120
      Width           =   2490
      Begin VB.OptionButton optNone 
         Caption         =   "Aucune"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.CheckBox chk20vrb 
         Caption         =   "sur 20 verbes"
         Enabled         =   0   'False
         Height          =   255
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.OptionButton optNoter 
         Caption         =   "Noter"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton optRedemande 
         Caption         =   "Redemande en cas d'erreur"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2280
      End
   End
   Begin VB.Frame fraMode 
      Caption         =   "Je vous donne :"
      Height          =   975
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   2295
      Begin VB.OptionButton optMode 
         Caption         =   "la traduction"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optMode 
         Caption         =   "un des 4 éléments"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1935
      End
   End
   Begin Threed.SSCommand SScmdAide 
      Height          =   405
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
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
      Picture         =   "Mode.frx":0614
   End
   Begin Threed.SSCommand SScmdAnnuler 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
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
      Picture         =   "Mode.frx":0886
   End
   Begin Threed.SSCommand SScmdOK 
      Default         =   -1  'True
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   2040
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
      Picture         =   "Mode.frx":0B38
   End
End
Attribute VB_Name = "frmMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAléatoire_Click()
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        MsgBox "Si cette case est cochée, le programme vous choisi un verbe aléatoirement parmi la liste en ne le faisant paraître que le nombre de fois nécessaire.", , "Aide sur 'ALÉATOIRE'"
        Exit Sub
    End If
End Sub

Private Sub cmdFichier_Click()
    On Error Resume Next

    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        MsgBox "Vous permet d'explorer votre disque dur pour changer de liste de verbes.", , "Aide sur 'FICHIER'"
        Exit Sub
    End If
    
    frmVIM!CommonDialog.ShowOpen
    If Err = cdlCancel Then Exit Sub
    
    Fichier = frmVIM!CommonDialog.filename
    txtFichier = Fichier
End Sub

Private Sub SScmdAide_Click()
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

Private Sub SScmdAnnuler_Click()
    ' Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Sort de la fenêtre pour revenir au menu principal.", , "Aide sur 'ANNULER'"
        Exit Sub
    End If
    SScmdOK.SetFocus
    
    Hide
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    
    txtFichier = ListVrb
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Test pour l'aide
    If KeyCode = 112 Then SScmdAide_Click
End Sub

Private Sub optMode_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Change le mode de questionnement:" & Enter & _
        "  Mode1: Donne le la traduction en français. Il faut trouver le reste." & Enter & _
        "  Mode2: Donne soit la traduction soit le verbe à l'indicatif soit" & Enter & _
        "               autre chose et il faut trouver les éléments manquants.", , "Aide sur 'MODE'"
    End If
End Sub

Private Sub optNoter_Click()
    ' Enlève/Met "sur 20 verbes"
    chk20vrb.Enabled = optNoter.Value
End Sub

Private Sub optNoter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Si vous cochez cette case, vous recevrez une note après l'interrogation." & Enter & _
          "Cela permet de simuler un contrôle." _
          , , "Aide sur 'NOTER'"
    End If
End Sub

Private Sub optRedemande_Click()
    ' Enlève/Met "sur 20 verbes"
    chk20vrb.Enabled = optNoter.Value
End Sub

Private Sub optRedemande_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Quand vous fait des fautes lors du questionnement," & Enter & _
          "le programme vous redemande ce(s) verbe(s) selon le nombre" & Enter & _
          "de fautes commises.", , "Aide sur 'REDEMANDE'"
    End If
End Sub

Private Sub SScmdOK_Click()
    On Error Resume Next
    
    ' Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Vous interroge selon le mode demandé.", , "Aide sur 'ANNULER'"
        Exit Sub
    End If
    
    ' Fichier existe ?
    If Dir(txtFichier) = "" Then
        MsgBox "Erreur: Spécifiez un nom de fichier valide", vbExclamation
        Exit Sub
    End If
    
    ListVrb = txtFichier
    
    Screen.MousePointer = 11
    SScmdAnnuler_Click
    If optMode(1) Then
        frmInterr1.Show 1
    Else
        frmInterr2.Show 1
    End If
End Sub
