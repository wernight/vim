VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprime une liste de verbes"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   11925
   Icon            =   "Print.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   15.478
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   21.034
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      DialogTitle     =   "Enregistrer sous"
      Filter          =   "*.* Tous les fichiers|*.*|*.TXT Texte|*.txt"
      FontName        =   "Courier New"
   End
   Begin VB.Line Ligne 
      Index           =   5
      X1              =   0
      X2              =   19.262
      Y1              =   1.693
      Y2              =   1.693
   End
   Begin VB.Line Ligne 
      Index           =   3
      X1              =   13.97
      X2              =   13.97
      Y1              =   1.27
      Y2              =   31.75
   End
   Begin VB.Line Ligne 
      Index           =   4
      X1              =   19.262
      X2              =   19.262
      Y1              =   1.27
      Y2              =   31.75
   End
   Begin VB.Line Ligne 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   1.27
      Y2              =   31.75
   End
   Begin VB.Label lblListeVrb 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   17970
      Index           =   3
      Left            =   8040
      TabIndex        =   8
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label lblListeVrb 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   18000
      Index           =   2
      Left            =   5400
      TabIndex        =   7
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblListeVrb 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   17970
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Line Ligne 
      Index           =   2
      X1              =   9.313
      X2              =   9.313
      Y1              =   1.27
      Y2              =   31.75
   End
   Begin VB.Line Ligne 
      Index           =   1
      X1              =   4.657
      X2              =   4.657
      Y1              =   1.27
      Y2              =   31.75
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Infinitif"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Traduction"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7920
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Participe passé"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Prétérit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label lblListeVrb 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   18000
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label lblTitre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liste de verbes irréguliers xxxxxxx."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4020
   End
   Begin VB.Shape Shape 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00000000&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   5  'Downward Diagonal
      Height          =   255
      Left            =   0
      Top             =   720
      Width           =   10935
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Dim Verbe As String, X As Integer, NbrPages As Integer, Lignes As Integer
Dim PrintSélec As Boolean, PrintFichier As Boolean, CopiesAssembl As Boolean
    On Error Resume Next
    Me.Top = 0
    Me.Left = 0
    
    ' Donne la langue
    For i% = 0 To 3
        If frmEdit!optLangue(i%) Then lblTitre = "Liste de verbes irréguliers " & Trim(frmEdit!optLangue(i%).Caption) & "."
    Next
    
    ' Config
    CommonDialog.Flags = cdlPDNoPageNums
    CommonDialog.ShowPrinter
    If Err = cdlCancel Then
        Hide
        Exit Sub
    End If
    
    ' Test le Boutons Séléctionnés ds CommenDialog
    PrintSélec = Not CommonDialog.Flags / 2 = Int(CommonDialog.Flags / 2)
    X = CommonDialog.Flags - cdlPDNoPageNums - cdlPDPrintToFile - (-PrintSélec)
    Select Case X
        Case cdlPDCollate
            PrintFichier = True
            CopiesAssembl = True
        Case 0
            PrintFichier = True
            CopiesAssembl = False
        Case -16
            PrintFichier = False
            CopiesAssembl = True
        Case Else
            PrintFichier = False
            CopiesAssembl = False
    End Select
    
    Rem Imprime
    For Copie% = 1 To CommonDialog.Copies
        If PrintFichier Then
            'Demande le Fichier
            CommonDialog.ShowSave
            If Err = cdlCancel Then
                Hide
                Exit Sub
            End If
            
            ' Enregistre
            Open CommonDialog.filename For Output As 1
            Print #1, "      " & lblTitre
            Print #1, "Infinitif      Prétérit       Participe      Traduction"
            'Met la liste
            For i% = 0 To frmEdit!lstVrb.ListCount - 2
                ' Si Print Sélection ET Verbe Séléectioné
                If PrintSélec = False Or frmEdit!Liste.Selected(i%) Then
                    Verbe = frmEdit!lstVrb.List(i%) & Enter
                    Verbe = Left(Verbe, Len(Verbe) - 1)
                    Print #1, Verbe
                End If
            Next
            Close
        Else
            ' Imprime
            For i% = 0 To frmEdit!lstVrb.ListCount - 2
                ' Si Print Sélection ET Verbe Séléectioné
                If PrintSélec = False Or frmEdit!Liste.Selected(i%) Then
                    Verbe = frmEdit!lstVrb.List(i%) & Enter
                    'Affiche
                    For j% = 0 To 3
                        lblListeVrb(j%) = lblListeVrb(j%) & Trim(Mid(Verbe, j% * 30 + 1, 30)) & Enter
                    Next
                    
                    Lignes = Lignes + 1
                End If
                
                'Ligne
                If i% = frmEdit!lstVrb.ListCount - 2 Then
                    h% = 14 * Lignes + 72
                    Ligne(5).Y1 = h%
                    For l% = 0 To 5
                        Ligne(l%).Y2 = h%
                        Ligne(l%).Y2 = h%
                    Next
                    PrintForm
                    Exit For
                End If
            
                'Si un lot de 37 lignes
                If Lignes >= 37 Then
                    PrintForm
                    'Efface
                    For c% = 0 To 3
                        lblListeVrb(c%) = ""
                    Next
                    Lignes = 0
                End If
            Next
        End If
    Next
    
    ' Sort
    Hide
    
    If Err Then MsgBox "Erreur: " & Error(Err) & "!", vbExclamation
End Sub
