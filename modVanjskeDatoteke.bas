Attribute VB_Name = "modVanjskeDatoteke"

Public Sub Spremi_INI_Datoteku(ImeDatoteke As String, Optional ShowMessage As Boolean)
Dim linija As String

    On Error GoTo ErrorHandler
    Open ImeDatoteke For Output Access Write As 1
        Print #1, frmGlavna.mnuViewProcess.Checked
        Print #1, frmGlavna.ComFilter.Caption
        Print #1, frmGlavna.txtXOrigin.Text
        Print #1, frmGlavna.txtYOrigin.Text
        Print #1, frmGlavna.txtXResolution.Text
        Print #1, frmGlavna.txtYResolution.Text
        Print #1, frmGlavna.txtMasterUnit.Text
        Print #1, frmGlavna.txtUorPerMaster.Text
        Print #1, frmGlavna.txtXGlobalOrigin.Text
        Print #1, frmGlavna.txtYGlobalOrigin.Text
        Print #1, frmGlavna.txtZGlobalOrigin.Text
        Print #1, frmGlavna.txtQuality.Text
        Print #1, frmGlavna.ChkTIFF.Value
        Print #1, frmGlavna.chkSIS.Value
        Print #1, frmGlavna.chkSJS.Value
        For i = 0 To frmGlavna.lstUlazneDatoteke.ListCount - 1
            Print #1, UCase(frmGlavna.lstUlazneDatoteke.List(i))
        Next i
        Close #1
            If ShowMessage Then
                a = MsgBox("File saved", vbInformation, "Information")
            End If
        Exit Sub
    
ErrorHandler:
    a = MsgBox("Pogreška prilikom otvaranja/rada s datotekom" + Chr(13) + ImeDatoteke, vbCritical, "Error")
    Close #1
End Sub
Public Sub Otvori_INI_Datoteku(ImeDatoteke As String)
Dim linija As String
    
    frmGlavna.lstUlazneDatoteke.Clear

    On Error GoTo ErrorHandler
    Open ImeDatoteke For Input Access Read As 1
        Input #1, linija
        frmGlavna.mnuViewProcess.Checked = linija
        Input #1, linija
        frmGlavna.File1.Pattern = linija
        Input #1, linija
        frmGlavna.txtXOrigin.Text = linija
        Input #1, linija
        frmGlavna.txtYOrigin.Text = linija
        Input #1, linija
        frmGlavna.txtXResolution.Text = linija
        Input #1, linija
        frmGlavna.txtYResolution.Text = linija
        Input #1, linija
        frmGlavna.txtMasterUnit.Text = linija
        Input #1, linija
        frmGlavna.txtUorPerMaster.Text = linija
        Input #1, linija
        frmGlavna.txtXGlobalOrigin.Text = linija
        Input #1, linija
        frmGlavna.txtYGlobalOrigin.Text = linija
        Input #1, linija
        frmGlavna.txtZGlobalOrigin.Text = linija
        Input #1, linija
        frmGlavna.txtQuality.Text = linija
        Input #1, linija
        frmGlavna.ChkTIFF.Value = linija
        Input #1, linija
        frmGlavna.chkSIS.Value = linija
        Input #1, linija
        frmGlavna.chkSJS.Value = linija
        Do While Not EOF(1)
            Input #1, linija
            frmGlavna.lstUlazneDatoteke.AddItem (linija)
        Loop
        Close #1
        frmGlavna.lblErrorMsg.Caption = ""
        frmGlavna.lblItemsCount.Caption = Str(frmGlavna.lstUlazneDatoteke.ListCount)
        Exit Sub
    
ErrorHandler:
    a = MsgBox("Can not open" + Chr(13) + ImeDatoteke, vbExclamation, "Error")
    Close #1
End Sub

Public Sub Otvori_TIFF2SDI_Datoteku(ImeDatoteke As String)

Dim Linija1 As String
Dim Linija2 As String
Dim Duzina As Integer
Dim Broj As Integer
Dim Commande(1 To 255) As String
    On Error GoTo ErrorHandler
    Open ImeDatoteke For Input Access Read As 1
    Line Input #1, Linija1
    Line Input #1, Linija2
    Close #1
    Duzina = 0
    Broj = 1
    Linija1 = Trim(Linija1) + " "
    Linija2 = Trim(Linija2) + " "
    For i = 1 To Len(Linija1) + 1
        If Mid(Linija1, i, 1) = " " Then
            Commande(Broj) = Trim(Mid(Linija1, i - Duzina, Duzina))
            Duzina = 0
            Broj = Broj + 1
        End If
    Duzina = Duzina + 1
    Next i
    Broj = 0
    For i = 1 To 255
        If Commande(i) <> "" Then
            Broj = Broj + 1
                Select Case Broj
                Case 1
                    frmGlavna.txtXOrigin.Text = Commande(i)
                Case 2
                    frmGlavna.txtYOrigin.Text = Commande(i)
                Case 3
                    frmGlavna.txtXResolution.Text = Commande(i)
                Case 4
                    frmGlavna.txtYResolution.Text = Commande(i)
                Case 5
                    frmGlavna.txtMasterUnit.Text = Commande(i)
                End Select
        End If
    Next i
    
    Broj = 1
    Duzina = 0
    For i = 1 To Len(Linija2) + 1
        If Mid(Linija2, i, 1) = " " Then
            Commande(Broj) = Trim(Mid(Linija2, i - Duzina, Duzina))
            Duzina = 0
            Broj = Broj + 1
        End If
    Duzina = Duzina + 1
    Next i
    Broj = 0
    For i = 1 To 255
        If Commande(i) <> "" Then
            Broj = Broj + 1
                Select Case Broj
                Case 1
                    frmGlavna.txtUorPerMaster.Text = Commande(i)
                Case 2
                    frmGlavna.txtXGlobalOrigin.Text = Commande(i)
                Case 3
                    frmGlavna.txtYGlobalOrigin.Text = Commande(i)
                Case 4
                    frmGlavna.txtZGlobalOrigin.Text = Commande(i)
                End Select
        End If
    Next i
    Exit Sub
    
    
ErrorHandler:
    a = MsgBox("Can not open" + Chr(13) + ImeDatoteke, vbExclamation, "Error")
    Close #1
End Sub

Public Sub Postavi_Defaultne_Vrijednosti()

frmGlavna.mnuViewProcess.Checked = False

frmGlavna.File1.Pattern = "*.tif"
frmGlavna.lblItemsCount.Caption = "0"
frmGlavna.txtXOrigin.Text = "0.0"
frmGlavna.txtYOrigin.Text = "0.0"
frmGlavna.txtXResolution.Text = "0.024"
frmGlavna.txtYResolution.Text = "0.024"
frmGlavna.txtMasterUnit.Text = "mm"
frmGlavna.txtUorPerMaster.Text = "10000.0"
frmGlavna.txtXGlobalOrigin.Text = "0.0"
frmGlavna.txtYGlobalOrigin.Text = "0.0"
frmGlavna.txtZGlobalOrigin.Text = "0.0"

frmGlavna.txtQuality.Text = "95"

frmGlavna.ChkTIFF.Value = 1
frmGlavna.chkSIS.Value = 1
frmGlavna.chkSJS.Value = 1


End Sub
Public Sub Spremi_TIFF2SDI_Datoteku(Optional ShowMessage As Boolean)
Dim ImeDatoteke As String

    ImeDatoteke = App.Path + "\TIFF2SDI.CFG"

    On Error Resume Next
        Kill (ImeDatoteke)
        
    On Error GoTo 0
    
    On Error GoTo ErrorHandler
        Open ImeDatoteke For Output Access Write As 1
        Print #1, Space(8 - Len(Trim(frmGlavna.txtXOrigin.Text))) + Trim(frmGlavna.txtXOrigin.Text) + Space(8 - Len(Trim(frmGlavna.txtYOrigin.Text))) + Trim(frmGlavna.txtYOrigin.Text) + Space(8 - Len(Trim(frmGlavna.txtXResolution.Text))) + Trim(frmGlavna.txtXResolution.Text) + Space(8 - Len(Trim(frmGlavna.txtYResolution.Text))) + Trim(frmGlavna.txtYResolution.Text) + Space(8) + Trim(frmGlavna.txtMasterUnit.Text)
        Print #1, Space(8 - Len(Trim(frmGlavna.txtUorPerMaster.Text))) + Trim(frmGlavna.txtUorPerMaster.Text) + Space(8 - Len(Trim(frmGlavna.txtXGlobalOrigin.Text))) + Trim(frmGlavna.txtXGlobalOrigin.Text) + Space(8 - Len(Trim(frmGlavna.txtYGlobalOrigin.Text))) + Trim(frmGlavna.txtYGlobalOrigin.Text) + Space(8 - Len(Trim(frmGlavna.txtZGlobalOrigin.Text))) + Trim(frmGlavna.txtZGlobalOrigin.Text)
        Close #1
            If ShowMessage Then
                a = MsgBox("File saved", vbInformation, "Information")
            End If
    Exit Sub
    
ErrorHandler:
    a = MsgBox("Can not save" + Chr(13) + ImeDatoteke, vbExclamation, "Error")
    Close #1
End Sub
