Attribute VB_Name = "Verbes"
Public Const PrixVIM As Integer = 6
Public Vrb() As String, FoisInterr() As Integer, NbrVrb As Integer, ListVrb As String
Public Elément(3, 1 To 16) As String
Public VrbEnCours As Integer, ttFautes As Integer
Public Reponse As String

Public Function ChargeListeVrb()
Dim Reponse As String, Verbe As String, Langue As Integer
Rem Charge la liste dans Vrb()
Rem FoisInterr() et NbrVrb
    On Error Resume Next
       
    ' Donne le rep. du prog.
    frmVIM!CommonDialog.InitDir = App.Path
    
    If Err = cdlCancel Then
        Err = 0
        DéjàFait = False
        ChargeListeVrb = "cdlCancel"
        Exit Function
    End If
    
    'Ouvre le fichier pour y prendre la liste des verbes
    Open ListVrb For Random As 1 Len = 122
    
    ' Affiche la langue
    Get #1, , Langue
    ChargeListeVrb = Langue
    
    ' Charge la liste des verbes
    NbrVrb = 0
    ttFautes = 0
    Do Until EOF(1)
        'Charge 1 ligne
        Get #1, , Verbe
        If Verbe = "" Then Exit Do
        NbrVrb = NbrVrb + 1
        ReDim Preserve Vrb(NbrVrb)
        Vrb(NbrVrb) = Verbe
    Loop
    Close
    DéjàFait = True
    VrbEnCours = 0
    ReDim FoisInterr(NbrVrb)
    For i% = 1 To NbrVrb
        FoisInterr(i%) = 1
    Next
End Function

Public Function VrbAléatoire()
Static NbrVrbInterr As Integer

    ' Test: Que sur 20 vrb
    If frmMode!optNoter And -frmMode!chk20vrb Then
        NbrVrbInterr = NbrVrbInterr + 1
        'Remet à 0
        If NbrVrbInterr = 22 Then NbrVrbInterr = 1
        'Test: 20 vrb max
        If NbrVrbInterr > 20 Then
            Call Noter
            GoTo Recommencer
        End If
    Else
        NbrVrbInterr = 0
    End If
    
    ' Test: si tous les verbes ont été faits
    For i% = 1 To NbrVrb
        If FoisInterr(i%) <> 0 Then
            'Génère un nombre aléatoire
            Randomize
            Do
                VrbEnCours = Int((NbrVrb * Rnd) + 1)
            Loop Until FoisInterr(VrbEnCours) <> 0
            FoisInterr(VrbEnCours) = FoisInterr(VrbEnCours) - 1
            Exit Function
        End If
    Next
    
    'Test: si Noter
    If frmMode!optNoter Then Call Noter
    
Recommencer:
    ' Si tout les vrb fait:
    If MsgBox("Fin des verbes." & Enter & "Recommencer?", vbYesNo) = vbYes Then
        VrbAléatoire = "Recommencer"
    Else
        ' Quitte
        DéjàFait = False
        VrbAléatoire = "Pas Recommencer"
    End If
End Function

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
    ' Test: VrbEnCours entier
    If Int(Note) = Note Then
        Msg$ = Msg$ & Format(Note, "#0") & "/20"
    Else
        Msg$ = Msg$ & Format(Note, "#0.#") & "/20"
    End If
    MsgBox Msg$
    
    ttFautes = 0
End Sub

