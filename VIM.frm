VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVIM 
   BackColor       =   &H00FF0000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verbes Irréguliers Multilingue"
   ClientHeight    =   5955
   ClientLeft      =   675
   ClientTop       =   1440
   ClientWidth     =   7965
   Icon            =   "VIM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MouseIcon       =   "VIM.frx":030A
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5955
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBulleAide 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   480
      Top             =   75
   End
   Begin VB.Timer Minuterie 
      Interval        =   1200
      Left            =   5070
      Top             =   -180
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.vrb"
      DialogTitle     =   "Ouvrir une liste de verbes"
      Filter          =   "Fichiers Verbe (*.vrb)|*.vrb|Tous les fichiers (*.*)|*.*"
   End
   Begin VB.Image imgRéviserUp 
      Height          =   315
      Left            =   3690
      Picture         =   "VIM.frx":0614
      Stretch         =   -1  'True
      Top             =   990
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image imgRéviserDn 
      Height          =   300
      Left            =   3690
      Picture         =   "VIM.frx":2028
      Stretch         =   -1  'True
      Top             =   1305
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgInterrUp 
      Height          =   375
      Left            =   4995
      Picture         =   "VIM.frx":3481
      Stretch         =   -1  'True
      Top             =   2205
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.Image imgInterrDn 
      Height          =   360
      Left            =   4995
      Picture         =   "VIM.frx":5420
      Stretch         =   -1  'True
      Top             =   2550
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Image imgEditUp 
      Height          =   495
      Left            =   6960
      Picture         =   "VIM.frx":6CE8
      Stretch         =   -1  'True
      Top             =   3330
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image imgEditDn 
      Height          =   465
      Left            =   6960
      Picture         =   "VIM.frx":99E1
      Stretch         =   -1  'True
      Top             =   3840
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Label lblBulleAide 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Msg d'aide."
      Height          =   195
      Left            =   1995
      TabIndex        =   0
      Top             =   4125
      Visible         =   0   'False
      Width           =   1770
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgApropos 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   5310
      Picture         =   "VIM.frx":BDF6
      Tag             =   "Informations sur l'auteur."
      Top             =   5235
      Width           =   2640
   End
   Begin VB.Image imgAproposUp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   5310
      Picture         =   "VIM.frx":11A54
      Top             =   5130
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image imgAproposDn 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   5310
      Picture         =   "VIM.frx":176B2
      Top             =   5055
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image imgAide 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   2670
      Picture         =   "VIM.frx":1D310
      Tag             =   "Présentation."
      Top             =   5235
      Width           =   2640
   End
   Begin VB.Image imgAideUp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   2670
      Picture         =   "VIM.frx":22F6E
      Top             =   5130
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image imgAideDn 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   2670
      Picture         =   "VIM.frx":28BCC
      Top             =   5055
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Image imgQuitter 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   15
      Picture         =   "VIM.frx":2E82A
      Tag             =   "Sortir du programme."
      Top             =   5235
      Width           =   2655
   End
   Begin VB.Image imgQuitterUp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   15
      Picture         =   "VIM.frx":3453C
      Top             =   5145
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgQuitterDn 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   705
      Left            =   15
      Picture         =   "VIM.frx":3A24E
      Top             =   5055
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Image imgRéviser 
      Height          =   795
      Left            =   720
      Picture         =   "VIM.frx":3FF60
      Top             =   810
      Width           =   2880
   End
   Begin VB.Image imgEdit 
      Height          =   975
      Left            =   720
      Picture         =   "VIM.frx":41974
      Top             =   3090
      Width           =   6240
   End
   Begin VB.Image imgInterr 
      Height          =   855
      Left            =   720
      Picture         =   "VIM.frx":4466D
      Top             =   2130
      Width           =   4290
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   8
      Left            =   4590
      Picture         =   "VIM.frx":4660C
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   7
      Left            =   4110
      Picture         =   "VIM.frx":46916
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   5
      Left            =   3150
      Picture         =   "VIM.frx":46D58
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   3
      Left            =   2190
      Picture         =   "VIM.frx":4719A
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   4
      Left            =   2670
      Picture         =   "VIM.frx":475DC
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   6
      Left            =   3630
      Picture         =   "VIM.frx":47A1E
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   2
      Left            =   1710
      Picture         =   "VIM.frx":47E60
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image icoVIM 
      Height          =   480
      Index           =   1
      Left            =   1230
      Picture         =   "VIM.frx":482A2
      Top             =   -180
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgBulleAide 
      Height          =   1050
      Left            =   1950
      Picture         =   "VIM.frx":485AC
      Top             =   4050
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Image imgFont 
      Height          =   4500
      Left            =   1815
      Picture         =   "VIM.frx":48B06
      Top             =   315
      Width           =   4500
   End
End
Attribute VB_Name = "frmVIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msgIntrouvable As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Test pour l'aide
    If KeyCode = 65 And Shift = vbAltMask Then imgAide_Click
    If KeyCode = 27 And Shift = 0 Then imgQuitter_Click
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrBulleAide_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "VIM", "Dernières listes", "Fichier", ListVrb
End Sub

Private Sub imgEdit_Click()
Dim Reponse As String, Fichier As String
    On Error Resume Next
    
    ' Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Vous permet de modifier ou d'imprimer une liste de verbes.", , "Aide sur 'LISTE DES VERBES'"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    frmEdit.Show
    Screen.MousePointer = 0
End Sub

Private Sub imgInterr_Click()
    'Test pour l'aide
    If Me.MousePointer <> 0 Then
        Me.MousePointer = 0
        MsgBox "Commence à vous interroger.", , "Aide sur 'INTERROGATION'"
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    frmMode.Show 1
End Sub

Private Sub imgRéviser_Click()
    Screen.MousePointer = 11
    frmModeRévision.Show 1
    Screen.MousePointer = 0
End Sub

Private Sub imgAide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgAide.Picture = imgAideDn.Picture
End Sub

Private Sub imgEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEdit = imgEditDn
End Sub

Private Sub imgEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgEdit = imgEditUp
End Sub

Private Sub imgInterr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgInterr = imgInterrDn
End Sub

Private Sub imginterr_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgInterr = imgInterrUp
End Sub

Private Sub imgRéviser_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRéviser = imgRéviserDn
End Sub

Private Sub imgRéviser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgRéviser = imgRéviserUp
End Sub

Private Sub Minuterie_Timer()
Static Index As Integer
Static FontIndex As Integer
    ' Icone
    Index = Index + 1: If Index = 9 Then Index = 1
    Me.Icon = icoVIM(Index)
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = 0
    Me.MousePointer = 0
End Sub

Private Sub Form_Load()
    AutoRedraw = True
    
    ' Centre le font
    imgFont.Move (Width - imgFont.Width) / 2, _
        (Height - imgQuitter.Height - imgFont.Height) / 2 - 50
    
    ListVrb = GetSetting("VIM", "Dernières listes", "Fichier")
    
    msgIntrouvable = Enter & "Fichier introuvable!" & Enter & "Voulez-vous continuer?"
    
    ' Donne le bon rep pour CommonDialog
    CommonDialog.Flags = cdlOFNHideReadOnly
    CommonDialog.InitDir = App.Path
    
    AutoRedraw = False
End Sub

Public Sub imgAide_Click()
    imgAide.Refresh

    Msg$ = "VIM est un logiciel éducatif qui vous permet de réviser vos verbes irréguliers en plusieurs langues." & Enter & _
        "Il vous suffit de choisir un fichier *.VRB (valide) dans un de vos disques pour charger une liste de verbes." & Enter & "Par exemple: le fichier ""2°_Ang.vrb"" est pour une classe de seconde en anglais première langue."
    MsgBox Msg$, , "Présentation du logiciel & Fichier VRB"
End Sub

Private Sub imgAide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgAide.Width Or Y < 0 Or Y > imgAide.Height Then
            imgAide.Picture = imgAideUp.Picture
        Else
            imgAide.Picture = imgAideDn.Picture
        End If
    End Select
    
    ' Bulle d'aide
    lblBulleAide.Caption = imgAide.Tag
    lblBulleAide.WordWrap = True
    lblBulleAide.Move imgAide.Left + X, imgAide.Top + Y + 280, 1770
    If lblBulleAide.Height < 200 Then lblBulleAide.WordWrap = False
    
    lblBulleAide.Visible = True
    PaintPicture imgBulleAide, lblBulleAide.Left, lblBulleAide.Top, lblBulleAide.Width, lblBulleAide.Height, 0, 0, lblBulleAide.Width, lblBulleAide.Height, vbSrcAnd
    
    tmrBulleAide.Enabled = True
End Sub

Private Sub imgAide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgAide.Picture = imgAideUp.Picture
End Sub

Private Sub imgAPropos_Click()
    imgApropos.Refresh

    Screen.MousePointer = 11
    frmAPropos.Show 1
End Sub

Private Sub imgAPropos_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgApropos.Picture = imgAproposDn.Picture
End Sub

Private Sub imgAPropos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgApropos.Width Or Y < 0 Or Y > imgApropos.Height Then
            imgApropos.Picture = imgAproposUp.Picture
        Else
            imgApropos.Picture = imgAproposDn.Picture
        End If
    End Select
    
    ' Bulle d'aide
    lblBulleAide.Caption = imgApropos.Tag
    lblBulleAide.WordWrap = True
    lblBulleAide.Move imgApropos.Left + X, imgApropos.Top + Y + 280, 1770
    If lblBulleAide.Height < 200 Then lblBulleAide.WordWrap = False
    
    lblBulleAide.Visible = True
    PaintPicture imgBulleAide, lblBulleAide.Left, lblBulleAide.Top, lblBulleAide.Width, lblBulleAide.Height, 0, 0, lblBulleAide.Width, lblBulleAide.Height, vbSrcAnd
    
    tmrBulleAide.Enabled = True
End Sub

Private Sub imgAPropos_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgApropos.Picture = imgAproposUp.Picture
End Sub

Private Sub imgQuitter_Click()
    imgQuitter.Refresh
    
    Call Form_Unload(0)
    End
End Sub

Private Sub imgQuitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgQuitter.Picture = imgQuitterDn.Picture
End Sub

Private Sub imgQuitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Si le bouton est activé, affiche le dessin supérieur lorsque la souris glisse
    ' hors de la zone du bouton ; sinon affiche le dessin inférieur.
    Select Case Button
    Case 1
        If X <= 0 Or X > imgQuitter.Width Or Y < 0 Or Y > imgQuitter.Height Then
            imgQuitter.Picture = imgQuitterUp.Picture
        Else
            imgQuitter.Picture = imgQuitterDn.Picture
        End If
    End Select
    
    ' Bulle d'aide
    lblBulleAide.Caption = imgQuitter.Tag
    lblBulleAide.WordWrap = True
    lblBulleAide.Move imgQuitter.Left + X, imgQuitter.Top + Y + 280, 1770
    If lblBulleAide.Height < 200 Then lblBulleAide.WordWrap = False
    
    lblBulleAide.Visible = True
    PaintPicture imgBulleAide, lblBulleAide.Left, lblBulleAide.Top, lblBulleAide.Width, lblBulleAide.Height, 0, 0, lblBulleAide.Width, lblBulleAide.Height, vbSrcAnd
    
    tmrBulleAide.Enabled = True
End Sub

Private Sub imgQuitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then imgQuitter.Picture = imgQuitterUp.Picture
End Sub

Private Sub tmrBulleAide_Timer()
    lblBulleAide.Visible = False
    tmrBulleAide.Enabled = False
End Sub
