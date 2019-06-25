Attribute VB_Name = "modGenerator"
Private tShellAndWait As udtShellAndWait    'store settings in the form in this structure

Public Sub Pokreni()
Dim listaSIS As String
Dim listaSJS As String
Dim listaDELTIFF As String
Dim listaDELSIS As String
Dim listaDELSJS As String
Dim BrojUListi As Integer
Dim i As Integer
Dim Pocetak, Vrijeme As Date

On Error Resume Next
    Kill (App.Path + "\TIFF2SIS2SJS.LOG")
On Error GoTo 0

Call modVanjskeDatoteke.Spremi_INI_Datoteku(App.Path + "\TIFF2SDI.INI", False)
Call modVanjskeDatoteke.Spremi_TIFF2SDI_Datoteku(False)
Call Blokiraj

Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
Print #1, Date$ + " " + Time$ + " - Process started"
Close #1

BrojUListi = frmGlavna.lstUlazneDatoteke.ListCount
i = 0

tShellAndWait.tStart.dwFlags = &H1

Do While frmGlavna.lstUlazneDatoteke.ListCount <> 0

    listaSIS = "T2s.bat " + Chr(34) + frmGlavna.lstUlazneDatoteke.List(0) + Chr(34)
    listaSJS = "s2j.bat " + Chr(34) + Mid$(frmGlavna.lstUlazneDatoteke.List(0), 1, Len(frmGlavna.lstUlazneDatoteke.List(0)) - 3) + "sis" + Chr(34) + " " + Chr(34) + Mid$(frmGlavna.lstUlazneDatoteke.List(0), 1, Len(frmGlavna.lstUlazneDatoteke.List(0)) - 3) + "sjs" + Chr(34) + " " + Chr(34) + frmGlavna.txtQuality
    listaDELTIFF = frmGlavna.lstUlazneDatoteke.List(0)
    listaDELSIS = Mid$(frmGlavna.lstUlazneDatoteke.List(0), 1, Len(frmGlavna.lstUlazneDatoteke.List(0)) - 3) + "sis"
    listaDELSJS = Mid$(frmGlavna.lstUlazneDatoteke.List(0), 1, Len(frmGlavna.lstUlazneDatoteke.List(0)) - 3) + "sjs"
        
    frmGlavna.lblNow.Caption = frmGlavna.lstUlazneDatoteke.List(0)
        
    If frmGlavna.mnuViewProcess.Checked = False Then
        tShellAndWait.tStart.wShowWindow = vbHide
    Else
        tShellAndWait.tStart.wShowWindow = vbNormalNoFocus
    End If
    
    Pocetak = Time
    
    tShellAndWait.sCommand = listaSIS
    ShellAndWait tShellAndWait
    tShellAndWait.sCommand = listaSJS
    ShellAndWait tShellAndWait
    
    
    If frmGlavna.ChkTIFF = 0 Then
        Call DeleteTIFF(listaDELTIFF)
    End If

    If frmGlavna.chkSIS = 0 Then
        Call DeleteSIS(listaDELSIS)
    End If

    If frmGlavna.chkSJS = 0 Then
        Call DeleteSJS(listaDELSJS)
    End If
    
    Vrijeme = Time - Pocetak
    
    Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
    Print #1, Date$ + " " + Time$ + " - " + Str(Vrijeme) + " - finished: " + frmGlavna.lstUlazneDatoteke.List(0)
    Close #1

    If frmGlavna.comStart.Enabled = False Then
        Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
        Print #1, Date$ + " " + Time$ + " - User stopped! "
        Close #1
        
        frmGlavna.lstUlazneDatoteke.RemoveItem (0)
        frmGlavna.lblItemsCount.Caption = Str(frmGlavna.lstUlazneDatoteke.ListCount)
        
        If MsgBox("Save current position?" + Chr(13) + "NOTE: If you save you will be able to continue your job later", vbYesNo, "Question") = vbYes Then
            Call Spremi_INI_Datoteku(App.Path + "\TIFF2SDI.INI", True)
        End If
        
        Call Odblokiraj
        Exit Sub
    End If
    
    i = i + 1
    frmGlavna.ProgressBar1.Value = (i / BrojUListi)
    frmGlavna.Caption = "TIFF2SIS2SJS... " + Str(CInt((i / BrojUListi) * 100)) + " %"
    
    frmGlavna.lstUlazneDatoteke.RemoveItem (0)
    frmGlavna.lblItemsCount.Caption = Str(frmGlavna.lstUlazneDatoteke.ListCount)

Loop

Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
Print #1, Date$ + " " + Time$ + " - Process finished"
Close #1

Call Odblokiraj
a = MsgBox("All files processed", vbInformation)
End Sub

Private Sub Blokiraj()

frmGlavna.comStart.Caption = "Stop"
frmGlavna.Drive1.Enabled = False
frmGlavna.Dir1.Enabled = Flase
frmGlavna.File1.Enabled = False
frmGlavna.ComFilter.Enabled = False
frmGlavna.mnuFile.Enabled = False
frmGlavna.lstUlazneDatoteke.Enabled = False
frmGlavna.comPrebaci.Enabled = False
frmGlavna.comPrebaciSve.Enabled = False
frmGlavna.comObrisi.Enabled = False
frmGlavna.comSaveConFile.Enabled = False
frmGlavna.comDefault95.Enabled = False
frmGlavna.txtQuality.Enabled = False
frmGlavna.ChkTIFF.Enabled = False
frmGlavna.chkSIS.Enabled = False
frmGlavna.chkSJS.Enabled = False

End Sub
Private Sub Odblokiraj()

frmGlavna.comStart.Caption = "Start"
frmGlavna.comStart.Enabled = True
frmGlavna.Drive1.Enabled = True
frmGlavna.Dir1.Enabled = True
frmGlavna.File1.Enabled = True
frmGlavna.ComFilter.Enabled = True
frmGlavna.mnuFile.Enabled = True
frmGlavna.lstUlazneDatoteke.Enabled = True
frmGlavna.comStart.Enabled = True
frmGlavna.comPrebaci.Enabled = True
frmGlavna.comPrebaciSve.Enabled = True
frmGlavna.comObrisi.Enabled = True
frmGlavna.comSaveConFile.Enabled = True
frmGlavna.comDefault95.Enabled = True
frmGlavna.txtQuality.Enabled = True
frmGlavna.ChkTIFF.Enabled = True
frmGlavna.chkSIS.Enabled = True
frmGlavna.chkSJS.Enabled = True
frmGlavna.lblNow.Caption = ""
frmGlavna.ProgressBar1.Value = 0
frmGlavna.Caption = "Tiff..2..Sis..2..Sjs"

End Sub

Private Sub DeleteTIFF(TIFFFileName As String)

On Error GoTo ErrorHandler
    Kill (TIFFFileName)
Exit Sub

ErrorHandler:
    Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
    Print #1, Date$ + " " + Time$ + " - ERROR - Can't delete file: " + TIFFFileName
    Close #1
    frmGlavna.lblErrorMsg.Caption = "You have error message(s) in LOG file!"
End Sub

Private Sub DeleteSIS(SISFileName As String)

On Error GoTo ErrorHandler
    Kill (SISFileName)
Exit Sub

ErrorHandler:
    Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
    Print #1, Date$ + " " + Time$ + " - ERROR - Can't delete file: " + SISFileName
    Close #1
    frmGlavna.lblErrorMsg.Caption = "You have error message(s) in LOG file!"
End Sub


Private Sub DeleteSJS(SJSFileName As String)

On Error GoTo ErrorHandler
    Kill (SJSFileName)
Exit Sub

ErrorHandler:
    Open App.Path + "\TIFF2SIS2SJS.LOG" For Append Access Write As 1
    Print #1, Date$ + " " + Time$ + " - ERROR - Can't delete file: " + SJSFileName
    Close #1
    frmGlavna.lblErrorMsg.Caption = "You have error message(s) in LOG file!"
End Sub
