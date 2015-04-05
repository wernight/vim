VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmModeRévision 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mode de révision"
   ClientHeight    =   1695
   ClientLeft      =   1590
   ClientTop       =   3165
   ClientWidth     =   5220
   Icon            =   "ModeRevi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "ModeRevi.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1695
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAléatoire 
      Caption         =   "Aléatoirement"
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.TextBox txtFichier 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   4080
   End
   Begin VB.CommandButton cmdFichier 
      Caption         =   "&Fichier"
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin Threed.SSCommand SScmdOK 
      Default         =   -1  'True
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   1080
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
      Picture         =   "ModeRevi.frx":0614
   End
   Begin Threed.SSCommand SScmdAnnuler 
      Cancel          =   -1  'True
      Height          =   405
      Left            =   1680
      TabIndex        =   1
      Top             =   1080
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
      Picture         =   "ModeRevi.frx":0C72
   End
   Begin Threed.SSCommand SScmdAide 
      Height          =   405
      Left            =   3600
      TabIndex        =   2
      Top             =   1080
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
      Picture         =   "ModeRevi.frx":0F24
   End
End
Attribute VB_Name = "frmModeRévision"
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

Private Sub Form_Activate()
    Screen.MousePointer = 0
    
    txtFichier = ListVrb
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Test pour l'aide
    If KeyCode = 112 Then SScmdAide_Click
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

Private Sub SScmdOK_Click()
    On Error Resume Next
    
    ' Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Vous fait réviser selon le mode demandé.", , "Aide sur 'ANNULER'"
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
    
    frmRéviser.Show 1
End Sub


