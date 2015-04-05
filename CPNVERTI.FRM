VERSION 4.00
Begin VB.Form Form1 
   Caption         =   "Converti de TXT à VRB"
   ClientHeight    =   1800
   ClientLeft      =   1560
   ClientTop       =   2310
   ClientWidth     =   5115
   Height          =   2205
   Left            =   1500
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   5115
   Top             =   1965
   Width           =   5235
   Begin VB.CommandButton cmdConvertir 
      Caption         =   "Convertir"
      Default         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1185
      TabIndex        =   2
      Top             =   1080
      Width           =   2640
   End
   Begin VB.TextBox txtFichier 
      Height          =   285
      Left            =   390
      TabIndex        =   1
      Text            =   "C:\All.txt"
      Top             =   390
      Width           =   3210
   End
   Begin VB.CommandButton cmdFichier 
      Caption         =   "Fichier"
      Height          =   270
      Left            =   3660
      TabIndex        =   0
      Top             =   390
      Width           =   1020
   End
   Begin VB.Label lblNbrVrb 
      Alignment       =   1  'Right Justify
      Caption         =   "xxx"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2430
      TabIndex        =   4
      Top             =   750
      Width           =   510
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre de verbes :"
      Height          =   210
      Left            =   870
      TabIndex        =   3
      Top             =   750
      Width           =   1500
   End
   Begin MSComDlg.CommonDialog cdFichier 
      Left            =   -15
      Top             =   -30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      DefaultExt      =   "*.TXT"
      Filter          =   "Fichiers texte (*.txt)|*.txt|Tous les fichier (*.*)|*.*"
      InitDir         =   "C:\"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdConvertir_Click()
Dim x As String, Vrb(1 To 4) As String, Verbe As String, Langue As Integer
    Screen.MousePointer = 11
    
    a% = 1
    Kill "C:\VIM.VRB"
    
    ' Prend les donnés
    Open txtFichier For Input As 1
    Open "C:\VIM.VRB" For Random As 2 Len = 122
    Langue = 1
    Put #2, , Langue
    Do Until EOF(1)
    
    i% = i% + 1
    
    Line Input #1, x
    'Les traites
    Do
        b% = InStr(a%, x, ", ") - 1
        If b% <> True Then
            n% = n% + 1
            Vrb(n%) = Mid(x, a%, b% - a% + 1)
            a% = b% + 3
        Else
            For n% = 1 To 3
                ' Saute l'élément 2
                If n% <> 2 Then Verbe = Verbe & Vrb(n%) & Space(30 - Len(Vrb(n%)))
            Next
            Verbe = Verbe & Mid(x, a%)
            Put #2, , Verbe
            Verbe = ""
            n% = 0
            a% = 1
            Exit Do
        End If
    Loop
    
    Loop
    Close
    
    lblNbrVrb = i%
    Screen.MousePointer = 0
End Sub
Private Sub cmdFichier_Click()
    cdFichier.ShowOpen
    txtFichier = cdFichier.filename
End Sub
