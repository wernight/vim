VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCopyRight 
   Caption         =   "VIM - Version COMPLETE -"
   ClientHeight    =   4275
   ClientLeft      =   1815
   ClientTop       =   2265
   ClientWidth     =   6645
   Icon            =   "CopyRigh.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4275
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSCommand cmdD�cliner 
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "D�cliner"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmdAccepter 
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Accepter"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCopyRight 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "xxx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmCopyRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAccepter_Click()
    Hide
End Sub

Private Sub cmdD�cliner_Click()
    End
End Sub

Private Sub Form_Load()
    lblCopyRight = " Cl� d'activation correcte.  Merci de vous �tre enregistr�..." & Enter & Enter & _
        "Il est formellement interdit:" & Enter & _
        "    -De modifier le code" & Enter & "    -De faire des copies d'�cran ou de documents" & Enter & _
        "    -De copier partiellement ou totalement le programme sauf dans sa version d'�valuation et une copie aux fins de s�curit�" & Enter & Enter & _
        "L'utilisation du programme sur une machine fixe et portable, appartenant au m�me utilisateur est autoris�e." & Enter & Enter & _
        "Vous �tes uniquement propri�taire du/des disques sur lesquels le titre est enregistr�, mais l'auteur est seul propri�taires du titre et de toutes les donn�es du titre." & Enter & Enter & _
        "Pour toute utilisation � titre professionnel et commercial, veuillez contacter l'auteur (voir 'A propos')."
End Sub
