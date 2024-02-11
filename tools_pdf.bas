Attribute VB_Name = "tools_pdf"
Option Explicit

Public Function Merge_Pdf(p_Form As Form, p_new_pdf As String, p_pagebg As String, Optional Calque As Boolean, Optional p_Societe As String) As String

    Dim Res     As Long
    Dim i       As Long
    Dim Tmp     As String
    Dim NbPages As Long
    Dim L_file  As String
    Dim Tmp2    As String
    On Error GoTo Err_Merge_Pdf
    
    Rem Le fichier temporaire est le fichier de création
    Tmp = Replace(p_new_pdf, ".pdf", "_tmp.pdf", , , vbTextCompare)
    Res = p_Form.PETOCX1.OpenOutputFile(Tmp)
    If Res <> 0 Then
        Merge_Pdf = "Impossible de créer le fichier résultat de la fusion PDF avec le fond de page"
        Exit Function
    End If
    
    If Calque Then
        NbPages = Read_Pdf_Num_Pages(p_Form, "", p_pagebg)
        L_file = p_pagebg
        If NbPages > 1 Then
            Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
        Else
            Select Case p_Societe
            'Case "033411000331", "033411004705"
            Case "033411000331"
            '    If p_Societe = "033411004705" Then
            '        res = p_Form.PETOCX1.AddLogo(p_new_pdf, 1)
            '    Else
                    Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
            '    End If
            Case Else
                Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
            End Select
        End If
    Else
        'Select Case p_Societe
        'Case "033411004705"
        '    res = p_Form.PETOCX1.AddLogo(p_pagebg, 0)
        'Case Else
            Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
        'End Select
    End If
    If Res <> 1 Then
        Rem Alors Problème!!!
        If Calque Then
            If NbPages > 1 Then
                Merge_Pdf = "Impossible de créer le fichier de destination"
            Else
                Merge_Pdf = "Impossible de lire le fichier joint"
            End If
        Else
            Merge_Pdf = "Impossible de lire le fichier de fond de page"
        End If
        While InStr(1, L_file, "\", vbTextCompare) > 0
            L_file = Mid(L_file, InStr(1, L_file, "\", vbTextCompare) + 1)
        Wend
        Merge_Pdf = Merge_Pdf & "(" & L_file & ")"
        Exit Function
    End If
        
    If Calque Then
        If NbPages > 1 Then
            If Pdf_ExtractPages(1, 1, p_pagebg, Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)) <> "Ok" Then
                Merge_Pdf = "Impossible d'extraire la première page du PDF (" & p_pagebg & ")"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            If Pdf_ExtractPages(2, NbPages, p_pagebg, Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)) <> "Ok" Then
                Merge_Pdf = "Impossible d'extraire les pages suivantes du PDF (" & p_pagebg & ")"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Res = p_Form.PETOCX1.MergeFile(Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare), 1, 0)
            If Res <= 0 Then
                Merge_Pdf = "Impossible de fusionner le fichier PDF avec le calque"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            p_Form.PETOCX1.CloseInputFile
            p_Form.PETOCX1.CloseOutputFile
            Kill p_new_pdf
Rem v508 - Fusion
            If Concat_PdfLib(p_Form, Tmp & "|" & Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare), p_new_pdf) <> "Ok" Then
                Merge_Pdf = "Impossible de fusionner le fichier PDF avec le calque (pages suivantes)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem Supprimer le fichier temporaire
            On Error Resume Next
            Kill Tmp
            Tmp2 = Replace(Tmp, "_tmp.pdf", "_main.pdf", , , vbTextCompare)
            If FileExists(Tmp2) Then
                Kill Tmp2
            End If
            Tmp2 = Replace(Tmp, "_tmp.pdf", "_main", , , vbTextCompare)
            If FileExists(Tmp2) Then
                Kill Tmp2
            End If
            Kill Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)
            Kill Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)
            Err.Clear
            Merge_Pdf = "Ok"
            Exit Function
        Else
            Rem 1 seule page
            Select Case p_Societe
            'Case "033411000331", "033411004705"
            Case "033411000331"
                Res = p_Form.PETOCX1.MergeFile(p_pagebg, 1, 0)
            Case Else
                Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
            End Select
        End If
    Else
        'stop
        Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
        'res = p_Form.PETOCX1.MergeFile(p_new_pdf, 1, 0)
    End If
    
    If Res <= 0 Then
        If Calque Then
            Merge_Pdf = "Impossible de fusionner le fichier PDF avec le calque"
        Else
            Merge_Pdf = "Impossible de fusionner le fichier PDF avec le fond de page"
        End If
        p_Form.PETOCX1.CloseOutputFile
        Exit Function
    End If
    
    
    p_Form.PETOCX1.CloseInputFile
    p_Form.PETOCX1.CloseOutputFile
    Rem Ecrase les fichier d'origine par le fichier créé
    FileCopy Tmp, p_new_pdf
    Rem Supprimer le fichier temporaire
    Kill Tmp
    
    Merge_Pdf = "Ok"
    Exit Function
    
Err_Merge_Pdf:
    Merge_Pdf = "KO - Merge Pdf - " & Err.Number & " - " & Err.Description
    
End Function

Public Function Merge_Pdf_New(p_Form As Form, p_new_pdf As String, p_pagebg As String, Optional Calque As Boolean, Optional p_Societe As String, Optional p_Out As String) As String

Rem v558 - Not used in fusion

    Dim Res     As Long
    Dim i       As Long
    Dim Tmp     As String
    Dim NbPages As Long
    Dim L_file  As String
    Dim Tmp2    As String
    
    On Error GoTo 0
    On Error GoTo Err_Merge_Pdf_New
    
    Rem Le fichier temporaire est le fichier de création
    Tmp = Replace(p_new_pdf, ".pdf", "_tmp.pdf", , , vbTextCompare)
    Res = p_Form.PETOCX1.OpenOutputFile(Tmp)
    If Res <> 0 Then
        Merge_Pdf_New = "Impossible de créer le fichier résultat de la fusion PDF avec le fond de page"
        Exit Function
    End If
    
    If Calque Then
        NbPages = Read_Pdf_Num_Pages(p_Form, "", p_pagebg)
        L_file = p_pagebg
        If NbPages > 1 Then
            Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
        Else
            Select Case p_Societe
            Case "033411000331", "033411004705"
            'Case "033411000331"
                
                If p_Societe = "033411004705" Then
                    Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
                Else
                    Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
                End If
            Case Else
                Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
            End Select
        End If
    Else
        'Select Case p_Societe
        'Case "033411004705"
        '    res = p_Form.PETOCX1.AddLogo(p_pagebg, 0)
        'Case Else
            Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
        'End Select
    End If
    If Res <> 1 Then
        If Calque Then
            If NbPages > 1 Then
                Merge_Pdf_New = "Impossible de créer le fichier de destination"
            Else
                Merge_Pdf_New = "Impossible de lire le fichier joint"
            End If
        Else
            Merge_Pdf_New = "Impossible de lire le fichier de fond de page"
        End If
        While InStr(1, L_file, "\", vbTextCompare) > 0
            L_file = Mid(L_file, InStr(1, L_file, "\", vbTextCompare) + 1)
        Wend
        Merge_Pdf_New = Merge_Pdf_New & "(" & L_file & ")"
        Exit Function
    End If
        
    If Calque Then
        If NbPages > 1 Then
            If Pdf_ExtractPages(1, 1, p_pagebg, Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)) <> "Ok" Then
                Merge_Pdf_New = "Impossible d'analyser le PDF joint (première page)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            If Pdf_ExtractPages(2, NbPages, p_pagebg, Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)) <> "Ok" Then
                Merge_Pdf_New = "Impossible d'analyser le PDF joint (pages suivantes)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Res = p_Form.PETOCX1.MergeFile(Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare), 1, 0)
            If Res <= 0 Then
                Merge_Pdf_New = "Impossible de fusionner le fichier PDF avec le calque"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            p_Form.PETOCX1.CloseInputFile
            p_Form.PETOCX1.CloseOutputFile
            Kill p_new_pdf
Rem v558
'stop
            'If Concat_Pdf(p_Form, Tmp & "|" & Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare), p_new_pdf) <> "Ok" Then
            If Concat_PdfLib(p_Form, Tmp & "|" & Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare), p_new_pdf) <> "Ok" Then
                Merge_Pdf_New = "Impossible de fusionner le fichier PDF avec le calque (pages suivantes)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem Supprimer le fichier temporaire
            On Error Resume Next
            Kill Tmp
            Tmp2 = Replace(Tmp, "_tmp.pdf", "_main.pdf", , , vbTextCompare)
            If FileExists(Tmp2) Then
                Kill Tmp2
            End If
            Tmp2 = Replace(Tmp, "_tmp.pdf", "_main", , , vbTextCompare)
            If FileExists(Tmp2) Then
                Kill Tmp2
            End If
            Kill Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)
            Kill Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)
            Err.Clear
            Merge_Pdf_New = "Ok"
            Exit Function
        Else
            Rem 1 seule page
            Select Case p_Societe
            Case "033411000331"
                Res = p_Form.PETOCX1.MergeFile(p_pagebg, 1, 0)
            Case Else
                Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
            End Select
        End If
    Else
        'stop
        Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
        'res = p_Form.PETOCX1.MergeFile(p_new_pdf, 1, 0)
    End If
    
    If Res <= 0 Then
        If Calque Then
            Merge_Pdf_New = "Impossible de fusionner le fichier PDF avec le calque"
        Else
            Merge_Pdf_New = "Impossible de fusionner le fichier PDF avec le fond de page"
        End If
        p_Form.PETOCX1.CloseOutputFile
        Exit Function
    End If
    
    
    p_Form.PETOCX1.CloseInputFile
    p_Form.PETOCX1.CloseOutputFile
    Rem Ecrase les fichier d'origine par le fichier créé
    If p_Out <> "" Then
        FileCopy Tmp, p_Out
    Else
        FileCopy Tmp, p_new_pdf
    End If
        Rem Supprimer le fichier temporaire
    Kill Tmp
        
    
    Merge_Pdf_New = "Ok"
    Exit Function
    
Err_Merge_Pdf_New:
    Merge_Pdf_New = "KO - Merge Pdf - " & Err.Number & " - " & Err.Description
    
End Function


Public Function Merge_Pdf_P2p(p_Form As Form, ByVal p_Pdf_In As String, p_Pdf_Fond As String, p_Pdf_Out As String) As String

    Dim Res     As Long
    Dim i       As Long
    Dim Result  As String
    Dim L_List  As String
    Dim L_Files_In() As String
    Dim L_Files_Out() As String
    Dim L_Files_Fond() As String
    Dim L_Nb_Pages  As Integer
    Dim L_In_Clean  As Boolean
    
    On Error GoTo 0
   
    Rem Compatibilité = Nombre de Pages
    If p_Form.PETOCX1.NumPages(p_Pdf_In) <> p_Form.PETOCX1.NumPages(p_Pdf_Fond) Then
        Merge_Pdf_P2p = "Nombre de pages différents (Fond de page : " & p_Form.PETOCX1.NumPages(p_Pdf_Fond) & " page(s) et document : " & p_Form.PETOCX1.NumPages(p_Pdf_In) & " page(s))"
        Exit Function
    End If
    L_Nb_Pages = p_Form.PETOCX1.NumPages(p_Pdf_In)
    
    Rem Controle
    If p_Pdf_In = p_Pdf_Out Then
        FileCopy p_Pdf_In, Replace(p_Pdf_In, ".pdf", "_in.pdf", 1, , vbTextCompare)
        p_Pdf_In = Replace(p_Pdf_In, ".pdf", "_in.pdf", 1, , vbTextCompare)
        L_In_Clean = True
    Else
        L_In_Clean = False
    End If
    
    Rem Alimentation variables
    ReDim L_Files_In(L_Nb_Pages + 1)
    ReDim L_Files_In(L_Nb_Pages + 1)
    ReDim L_Files_Fond(L_Nb_Pages + 1)
    ReDim L_Files_Out(L_Nb_Pages + 1)
    For i = 1 To L_Nb_Pages
        L_Files_Out(i) = Replace(p_Pdf_Out, ".pdf", Right("000" & i, 3) & ".pdf", , , vbTextCompare)
        L_Files_In(i) = Replace(p_Pdf_In, ".pdf", Right("000" & i, 3) & ".pdf", , , vbTextCompare)
        L_Files_Fond(i) = Replace(p_Pdf_Fond, ".pdf", Right("000" & i, 3) & ".pdf", , , vbTextCompare)
    Next
    
    Rem découper / fusionner page à page / concaténer
    For i = 1 To L_Nb_Pages
        Result = Pdf_ExtractPages(i, i, p_Pdf_In, L_Files_In(i))
        If Result <> "Ok" Then
            Merge_Pdf_P2p = Result
            GoTo Exit_Function
        End If
        Call Pdf_ExtractPages(i, i, p_Pdf_Fond, L_Files_Fond(i))
        If Result <> "Ok" Then
            Merge_Pdf_P2p = Result
            GoTo Exit_Function
        End If
    Next
    
    Rem Fusion page à page
    For i = 1 To L_Nb_Pages
    
        Res = p_Form.PETOCX1.OpenOutputFile(L_Files_Out(i))
        If Res <> 0 Then
            Merge_Pdf_P2p = p_Pdf_Out & "Impossible de créer le fichier"
            GoTo Exit_Function
        End If

        Res = p_Form.PETOCX1.AddLogo(L_Files_Fond(i), 0)
        If Res <> 1 Then
            Merge_Pdf_P2p = p_Pdf_Fond & "Impossible de créer le fichier"
            GoTo Exit_Function
        End If
        
        Res = p_Form.PETOCX1.MergeFile(L_Files_In(i), 0, 0)
        If Res <= 0 Then
            Merge_Pdf_P2p = p_Pdf_In & "Impossible de fusionner le fichier"
            p_Form.PETOCX1.CloseOutputFile
            GoTo Exit_Function
        End If
        p_Form.PETOCX1.CloseInputFile
        p_Form.PETOCX1.CloseOutputFile
    Next
    
    Rem Concaténation des pages générées
    L_List = ""
    For i = 1 To L_Nb_Pages
        L_List = L_List & L_Files_Out(i) & "|"
    Next
    L_List = Left(L_List, Len(L_List) - 1)
    
Rem v558
    'stop
    'Result = Concat_Pdf(p_Form, L_List, p_Pdf_Out)
    Result = Concat_PdfLib(p_Form, L_List, p_Pdf_Out)
    If Result <> "Ok" Then
        Merge_Pdf_P2p = "Impossible de concaténer le fichier de sortie (" & p_Pdf_Out & ")"
    End If
    
    
    'ShellExecute 0, "open", p_Pdf_Out, "", "", 1
    
    Rem Suppression / Nettoyage
Exit_Function:

    On Error Resume Next
    For i = 1 To L_Nb_Pages
        Kill L_Files_Out(i)
        Kill L_Files_In(i)
        Kill L_Files_Fond(i)
    Next
    
    If L_In_Clean Then
        Kill p_Pdf_In
    End If
    
    If Merge_Pdf_P2p = vbNullString Then
        Merge_Pdf_P2p = "Ok"
    End If
    

End Function

Public Function Merge_Pdf_P2p_New(p_Form As Form, ByVal p_Pdf_In As String, p_Pdf_Fond As String, p_Pdf_Out As String) As String

    Dim Res     As Long
    Dim i       As Long
    Dim Result  As String
    Dim L_List  As String
    Dim L_Files_In() As String
    Dim L_Files_Out() As String
    Dim L_Files_Fond() As String
    Dim L_Nb_Pages  As Integer
    Dim L_Nb_PagesMax   As Integer
    Dim L_In_Clean  As Boolean
    Dim L_ContactRest   As String
    
    On Error GoTo 0
   
    Rem Compatibilité = Nombre de Pages
    If p_Form.PETOCX1.NumPages(p_Pdf_In) < p_Form.PETOCX1.NumPages(p_Pdf_Fond) Then
        Merge_Pdf_P2p_New = "Nombre de pages différents (Fond de page : " & p_Form.PETOCX1.NumPages(p_Pdf_Fond) & " page(s) et document : " & p_Form.PETOCX1.NumPages(p_Pdf_In) & " page(s))"
        Exit Function
    End If
    L_Nb_Pages = p_Form.PETOCX1.NumPages(p_Pdf_In)
    L_Nb_PagesMax = p_Form.PETOCX1.NumPages(p_Pdf_Fond)
    L_ContactRest = ""
    Rem Controle
    If p_Pdf_In = p_Pdf_Out Then
        FileCopy p_Pdf_In, Replace(p_Pdf_In, ".pdf", "_in.pdf", 1, , vbTextCompare)
        p_Pdf_In = Replace(p_Pdf_In, ".pdf", "_in.pdf", 1, , vbTextCompare)
        L_In_Clean = True
    Else
        L_In_Clean = False
    End If
    
    Rem Alimentation variables
    ReDim L_Files_In(L_Nb_PagesMax + 1)
    ReDim L_Files_In(L_Nb_PagesMax + 1)
    ReDim L_Files_Fond(L_Nb_PagesMax + 1)
    ReDim L_Files_Out(L_Nb_PagesMax + 1)
    For i = 1 To L_Nb_PagesMax
        L_Files_Out(i) = Replace(p_Pdf_Out, ".pdf", "P" & Right("00" & i, 3) & ".pdf", , , vbTextCompare)
        L_Files_In(i) = Replace(p_Pdf_In, ".pdf", "P" & Right("00" & i, 3) & ".pdf", , , vbTextCompare)
        L_Files_Fond(i) = Replace(p_Pdf_Fond, ".pdf", "P" & Right("00" & i, 3) & ".pdf", , , vbTextCompare)
    Next
    
    Rem découper / fusionner page à page / concaténer
    For i = 1 To L_Nb_PagesMax
        Result = Pdf_ExtractPages(i, i, p_Pdf_In, L_Files_In(i))
        If Result <> "Ok" Then
            Merge_Pdf_P2p_New = Result
            GoTo Exit_Function
        End If
        Result = Pdf_ExtractPages(i, i, p_Pdf_Fond, L_Files_Fond(i), True)
        If Result <> "Ok" Then
            Merge_Pdf_P2p_New = Result
            GoTo Exit_Function
        End If
    Next
    
    Rem Fusion page à page
    For i = 1 To L_Nb_PagesMax
    
        Res = p_Form.PETOCX1.OpenOutputFile(L_Files_Out(i))
        If Res <> 0 Then
            Merge_Pdf_P2p_New = p_Pdf_Out & "Impossible de créer le fichier"
            GoTo Exit_Function
        End If

        Res = p_Form.PETOCX1.AddLogo(L_Files_Fond(i), 0)
        If Res <> 1 Then
            Merge_Pdf_P2p_New = p_Pdf_Fond & "Impossible de créer le fichier"
            GoTo Exit_Function
        End If
        
        Res = p_Form.PETOCX1.MergeFile(L_Files_In(i), 0, 0)
        If Res <= 0 Then
            Merge_Pdf_P2p_New = p_Pdf_In & "Impossible de fusionner le fichier"
            p_Form.PETOCX1.CloseOutputFile
            GoTo Exit_Function
        End If
        p_Form.PETOCX1.CloseInputFile
        p_Form.PETOCX1.CloseOutputFile
    Next
    
    If L_Nb_PagesMax < L_Nb_Pages Then
        L_ContactRest = Replace(p_Pdf_In, ".pdf", "rest.pdf", , , vbTextCompare)
        Result = Pdf_ExtractPages(L_Nb_PagesMax + 1, L_Nb_Pages + 0, p_Pdf_In, L_ContactRest)
        If Result <> "Ok" Then
            Merge_Pdf_P2p_New = Result
            GoTo Exit_Function
        End If
    End If
    Rem Concaténation des pages générées
    L_List = ""
    For i = 1 To L_Nb_PagesMax
        L_List = L_List & L_Files_Out(i) & "|"
    Next
    If L_ContactRest <> "" Then
        L_List = L_List & L_ContactRest
    Else
        L_List = Left(L_List, Len(L_List) - 1)
    End If
    
Rem v558
    'stop
    'Result = Concat_Pdf(p_Form, L_List, p_Pdf_Out)
    Result = Concat_PdfLib(p_Form, L_List, p_Pdf_Out)
    If Result <> "Ok" Then
        Merge_Pdf_P2p_New = "Impossible de concaténer le fichier de sortie (" & p_Pdf_Out & ")"
    End If
    'Exit Function
    
    'ShellExecute 0, "open", p_Pdf_Out, "", "", 1
    
    Rem Suppression / Nettoyage
Exit_Function:

    On Error Resume Next
    For i = 1 To L_Nb_PagesMax
        Kill L_Files_Out(i)
        Kill L_Files_In(i)
        Rem LCI pour CPAM le 26/03/2017
        'Kill L_Files_Fond(i)
    Next
    
    If L_In_Clean Then
        Kill p_Pdf_In
    End If
    
    If Merge_Pdf_P2p_New = vbNullString Then
        Merge_Pdf_P2p_New = "Ok"
    End If
    

End Function

Public Function Merge_Pdf_NewAuto(p_Form As Form, p_new_pdf As String, p_pagebg As String, Optional Calque As Boolean, Optional p_Societe As String, Optional p_Out As String) As String

Rem v558 - Not used in fusion

    Dim Res     As Long
    Dim Result  As String
    Dim i       As Long
    Dim Tmp     As String
    Dim NbPagesFond As Long
    Dim NbPagesDoc As Long
    
    Dim L_file  As String
    Dim tmp1    As String
    Dim Tmp2    As String
    Dim Tmp3    As String
    Dim Tmp4    As String
    Dim Tmp5    As String
    Dim Tmp1f   As String
    Dim Tmp2f   As String
    Dim Tmp3f   As String
    Dim Tmp4f   As String
    Dim Tmp5f   As String
    
    
    On Error GoTo 0
    On Error GoTo Err_Merge_Pdf_NewAuto
    
    Rem Pour gérer la fusion automatique, voici ce qui est retenu
    Rem Si le nombre de pages du fond = 1 alors, on fait comme d'habitude CAS 1
    
    Rem *************************************************************************************
    Rem *************************************************************************************
    Rem CAS 1
    Rem *************************************************************************************
    Rem *************************************************************************************
    NbPagesFond = Read_Pdf_Num_Pages(p_Form, "", p_pagebg)
    
    If NbPagesFond = 1 Then
        Rem COMME AVANT
    Rem Le fichier temporaire est le fichier de création
        Tmp = Replace(p_new_pdf, ".pdf", "_tmp.pdf", , , vbTextCompare)
        Res = p_Form.PETOCX1.OpenOutputFile(Tmp)
        If Res <> 0 Then
            Merge_Pdf_NewAuto = "Impossible de créer le fichier résultat de la fusion PDF avec le fond de page"
            Exit Function
        End If
        
        If Calque Then
            NbPagesFond = Read_Pdf_Num_Pages(p_Form, "", p_pagebg)
            L_file = p_pagebg
            If NbPagesFond > 1 Then
                Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
            Else
                Select Case p_Societe
                Case "033411000331", "033411004705"
                'Case "033411000331"
                    
                    If p_Societe = "033411004705" Then
                        Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
                    Else
                        Res = p_Form.PETOCX1.AddLogo(p_new_pdf, 0)
                    End If
                Case Else
                    Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
                End Select
            End If
        Else
            'Select Case p_Societe
            'Case "033411004705"
            '    res = p_Form.PETOCX1.AddLogo(p_pagebg, 0)
            'Case Else
                Res = p_Form.PETOCX1.AddLogo(p_pagebg, 1)
            'End Select
        End If
        If Res <> 1 Then
            If Calque Then
                If NbPagesFond > 1 Then
                    Merge_Pdf_NewAuto = "Impossible de créer le fichier de destination"
                Else
                    Merge_Pdf_NewAuto = "Impossible de lire le fichier joint"
                End If
            Else
                Merge_Pdf_NewAuto = "Impossible de lire le fichier de fond de page"
            End If
            While InStr(1, L_file, "\", vbTextCompare) > 0
                L_file = Mid(L_file, InStr(1, L_file, "\", vbTextCompare) + 1)
            Wend
            Merge_Pdf_NewAuto = Merge_Pdf_NewAuto & "(" & L_file & ")"
            Exit Function
        End If
            
        If Calque Then
            If NbPagesFond > 1 Then
                If Pdf_ExtractPages(1, 1, p_pagebg, Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)) <> "Ok" Then
                    Merge_Pdf_NewAuto = "Impossible d'analyser le PDF joint (première page)"
                    p_Form.PETOCX1.CloseOutputFile
                    Exit Function
                End If
                If Pdf_ExtractPages(2, NbPagesFond, p_pagebg, Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)) <> "Ok" Then
                    Merge_Pdf_NewAuto = "Impossible d'analyser le PDF joint (pages suivantes)"
                    p_Form.PETOCX1.CloseOutputFile
                    Exit Function
                End If
                Res = p_Form.PETOCX1.MergeFile(Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare), 1, 0)
                If Res <= 0 Then
                    Merge_Pdf_NewAuto = "Impossible de fusionner le fichier PDF avec le calque"
                    p_Form.PETOCX1.CloseOutputFile
                    Exit Function
                End If
                p_Form.PETOCX1.CloseInputFile
                p_Form.PETOCX1.CloseOutputFile
                Kill p_new_pdf
    Rem v558
    'stop
                'If Concat_Pdf(p_Form, Tmp & "|" & Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare), p_new_pdf) <> "Ok" Then
                If Concat_PdfLib(p_Form, Tmp & "|" & Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare), p_new_pdf) <> "Ok" Then
                    Merge_Pdf_NewAuto = "Impossible de fusionner le fichier PDF avec le calque (pages suivantes)"
                    p_Form.PETOCX1.CloseOutputFile
                    Exit Function
                End If
                Rem Supprimer le fichier temporaire
                On Error Resume Next
                Kill Tmp
                Tmp2 = Replace(Tmp, "_tmp.pdf", "_main.pdf", , , vbTextCompare)
                If FileExists(Tmp2) Then
                    Kill Tmp2
                End If
                Tmp2 = Replace(Tmp, "_tmp.pdf", "_main", , , vbTextCompare)
                If FileExists(Tmp2) Then
                    Kill Tmp2
                End If
                Kill Replace(p_pagebg, ".pdf", "_1.pdf", , , vbTextCompare)
                Kill Replace(p_pagebg, ".pdf", "_n.pdf", , , vbTextCompare)
                Err.Clear
                Merge_Pdf_NewAuto = "Ok"
                Exit Function
            Else
                Rem 1 seule page
                Select Case p_Societe
                Case "033411000331"
                    Res = p_Form.PETOCX1.MergeFile(p_pagebg, 1, 0)
                Case Else
                    Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
                End Select
            End If
        Else
            'stop
            Res = p_Form.PETOCX1.MergeFile(p_new_pdf, 0, 0)
            'res = p_Form.PETOCX1.MergeFile(p_new_pdf, 1, 0)
        End If
        
        If Res <= 0 Then
            If Calque Then
                Merge_Pdf_NewAuto = "Impossible de fusionner le fichier PDF avec le calque"
            Else
                Merge_Pdf_NewAuto = "Impossible de fusionner le fichier PDF avec le fond de page"
            End If
            p_Form.PETOCX1.CloseOutputFile
            Exit Function
        End If
        
        
        p_Form.PETOCX1.CloseInputFile
        p_Form.PETOCX1.CloseOutputFile
        Rem Ecrase les fichier d'origine par le fichier créé
        If p_Out <> "" Then
            FileCopy Tmp, p_Out
        Else
            FileCopy Tmp, p_new_pdf
        End If
            Rem Supprimer le fichier temporaire
        Kill Tmp
    Else
        Rem Dans ce cas, il y a au moins deux pages pour le fond, la règle est la suivante
        Rem Pour chaque page du fond, sauf la dernière, on concatène page à page,
        Rem puis on fusionne les pages restantes du document avec la dernière page
        NbPagesDoc = Read_Pdf_Num_Pages(p_Form, "", p_new_pdf)
        
        If NbPagesDoc <= NbPagesFond Then
            Rem On extrait les pages de fond dans un PDF, puis on merge page à page
            If NbPagesDoc < NbPagesFond Then
                Tmp1f = Replace(p_pagebg, ".pdf", "_1_" & NbPagesDoc & ".pdf", , , vbTextCompare)
                If Pdf_ExtractPages(1, NbPagesDoc, p_pagebg, Tmp1f) <> "Ok" Then
                    Merge_Pdf_NewAuto = "Impossible d'extraire les pages du fond (ERR:PSup01)"
                    p_Form.PETOCX1.CloseOutputFile
                    Exit Function
                End If
            Else 'Cas ou le nombre de page est égal
                Tmp1f = p_pagebg
            End If
            Rem à ce stade, il reste à merger page à page le document d'origine et le nouveau fond
            Result = Merge_Pdf_P2p_New(p_Form, p_new_pdf, Tmp1f, p_Out)
            If Result <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible de fusionner page à page (ERR:PSup02)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
        Else
            Rem Cas ou le nombre de pages du document est supérieur au nombre de pages du fond
            Rem On extrait NbPagesFond - 1, on merge P2P
            Rem On merge le reste du doc avec la dernière page du fond
            Rem On concatène les deux
            
            Rem On extrait les fond-1 premières pages
            tmp1 = Replace(p_new_pdf, ".pdf", "_1_" & NbPagesFond - 1 & ".pdf", , , vbTextCompare)
            If Pdf_ExtractPages(1, NbPagesFond - 1, p_new_pdf, tmp1) <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible d'extraire les pages du fond (ERR:PSup03)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Tmp1f = Replace(p_pagebg, ".pdf", "_1_" & NbPagesFond - 1 & ".pdf", , , vbTextCompare)
            If Pdf_ExtractPages(1, NbPagesFond - 1, p_pagebg, Tmp1f, True) <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible d'extraire les pages du fond (ERR:PSup04)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem On fusionne P à P les deux fichiers extraits
            'Result = Merge_Pdf_P2p_New(p_Form, Tmp1, Tmp1f, Tmp1)
            Result = Merge_Pdf_NewAuto(p_Form, tmp1, Tmp1f, Calque, p_Societe, tmp1)
            If Result <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible de fusionner page à page (ERR:PSup05)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem On extrait les dernières pages du doc
            Tmp2 = Replace(p_new_pdf, ".pdf", "_" & NbPagesFond & "_" & NbPagesDoc & ".pdf", , , vbTextCompare)
            If Pdf_ExtractPages(NbPagesFond, NbPagesDoc, p_new_pdf, Tmp2) <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible d'extraire les pages du fond (ERR:PSup05)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem On extrait la dernière page du fond
            Tmp2f = Replace(p_pagebg, ".pdf", "_" & NbPagesFond & "_" & NbPagesFond & ".pdf", , , vbTextCompare)
            If Pdf_ExtractPages(NbPagesFond, NbPagesFond, p_pagebg, Tmp2f, True) <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible d'extraire les pages du fond (ERR:PSup06)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem On fusionne les deux
            Result = Merge_Pdf_NewAuto(p_Form, Tmp2, Tmp2f, Calque, p_Societe, Tmp2)
            If Result <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible de fusionner page à page (ERR:PSup07)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
            Rem On concatène les deux fichier résultats
            If Concat_PdfLib(p_Form, tmp1 & "|" & Tmp2, p_Out) <> "Ok" Then
                Merge_Pdf_NewAuto = "Impossible de concaténer les fichiers ((ERR:PSup08)"
                p_Form.PETOCX1.CloseOutputFile
                Exit Function
            End If
        End If
    End If
    
    Merge_Pdf_NewAuto = "Ok"
    Exit Function
    
Err_Merge_Pdf_NewAuto:
    Merge_Pdf_NewAuto = "KO - Merge Pdf - " & Err.Number & " - " & Err.Description
    
End Function

