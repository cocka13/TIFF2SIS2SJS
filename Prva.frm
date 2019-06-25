VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmGlavna 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tiff..2..Sis..2..Sjs"
   ClientHeight    =   5175
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   8910
   Icon            =   "Prva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4920
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   8705
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Input Files"
      TabPicture(0)   =   "Prva.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label13"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblItemsCount"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Drive1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Dir1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "File1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "comPrebaci"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "comPrebaciSve"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "comObrisi"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ComFilter"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lstUlazneDatoteke"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "comObrisiSve"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Configuration File"
      TabPicture(1)   =   "Prva.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "txtZGlobalOrigin"
      Tab(1).Control(2)=   "txtYGlobalOrigin"
      Tab(1).Control(3)=   "txtXGlobalOrigin"
      Tab(1).Control(4)=   "txtUorPerMaster"
      Tab(1).Control(5)=   "txtMasterUnit"
      Tab(1).Control(6)=   "txtYResolution"
      Tab(1).Control(7)=   "txtXResolution"
      Tab(1).Control(8)=   "txtYOrigin"
      Tab(1).Control(9)=   "txtXOrigin"
      Tab(1).Control(10)=   "Frame1"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Process Monitor"
      TabPicture(2)   =   "Prva.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).Control(1)=   "Frame4"
      Tab(2).Control(2)=   "comStart"
      Tab(2).Control(3)=   "Frame3"
      Tab(2).ControlCount=   4
      Begin VB.CommandButton comObrisiSve 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   48
         ToolTipText     =   "Delete all files in List"
         Top             =   3840
         Width           =   615
      End
      Begin VB.Frame Frame5 
         Caption         =   "LOG File"
         Height          =   1215
         Left            =   -74880
         TabIndex        =   44
         Top             =   1560
         Width           =   8655
         Begin VB.CommandButton comLOGFile 
            Caption         =   "View LOG file"
            Height          =   255
            Left            =   6960
            TabIndex        =   47
            Top             =   600
            Width           =   1455
         End
         Begin VB.Label lblErrorMsg 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1440
            TabIndex        =   46
            Top             =   600
            Width           =   5355
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Error messages:"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   600
            Width           =   1125
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Process Monitor"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   41
         Top             =   2880
         Width           =   8655
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Message..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Width           =   900
         End
         Begin VB.Label lblNow 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1320
            TabIndex        =   42
            Top             =   480
            Width           =   7125
         End
      End
      Begin VB.CommandButton comStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74520
         TabIndex        =   40
         Top             =   4200
         Width           =   7935
      End
      Begin VB.Frame Frame3 
         Caption         =   "TIFF....SIS...SJS"
         Height          =   855
         Left            =   -74880
         TabIndex        =   34
         Top             =   600
         Width           =   8655
         Begin VB.CheckBox chkSJS 
            Caption         =   "SJS"
            Height          =   255
            Left            =   6240
            TabIndex        =   37
            ToolTipText     =   "If is not checked, SJS files will be deleted after process"
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chkSIS 
            Caption         =   "SIS"
            Height          =   255
            Left            =   3960
            TabIndex        =   36
            ToolTipText     =   "If is not checked, SIS files will be deleted after process"
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox ChkTIFF 
            Caption         =   "TIFF"
            Height          =   255
            Left            =   1800
            TabIndex        =   35
            ToolTipText     =   "If is not checked, TIFF files will be deleted after process"
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "----->"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5160
            TabIndex        =   39
            Top             =   360
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "----->"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2880
            TabIndex        =   38
            Top             =   360
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "SIS 2 SJS"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   30
         Top             =   3000
         Width           =   8655
         Begin VB.CommandButton comDefault95 
            Caption         =   "Default (95%)"
            Height          =   255
            Left            =   5280
            TabIndex        =   33
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txtQuality 
            Height          =   285
            Left            =   3600
            MaxLength       =   2
            TabIndex        =   31
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "SJS Quality:"
            Height          =   315
            Left            =   2160
            TabIndex        =   32
            Top             =   480
            Width           =   855
         End
      End
      Begin VB.TextBox txtZGlobalOrigin 
         Height          =   285
         Left            =   -69600
         MaxLength       =   8
         TabIndex        =   18
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtYGlobalOrigin 
         Height          =   285
         Left            =   -71280
         MaxLength       =   8
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtXGlobalOrigin 
         Height          =   285
         Left            =   -72960
         MaxLength       =   8
         TabIndex        =   16
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtUorPerMaster 
         Height          =   285
         Left            =   -74640
         MaxLength       =   8
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtMasterUnit 
         Height          =   285
         Left            =   -67920
         TabIndex        =   14
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtYResolution 
         Height          =   285
         Left            =   -69600
         MaxLength       =   8
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtXResolution 
         Height          =   285
         Left            =   -71280
         MaxLength       =   8
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtYOrigin 
         Height          =   285
         Left            =   -72960
         MaxLength       =   8
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtXOrigin 
         Height          =   285
         Left            =   -74640
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox lstUlazneDatoteke 
         Height          =   4155
         ItemData        =   "Prva.frx":68A6
         Left            =   3960
         List            =   "Prva.frx":68A8
         MultiSelect     =   2  'Extended
         TabIndex        =   9
         Top             =   360
         Width           =   4815
      End
      Begin VB.CommandButton ComFilter 
         Caption         =   "*.*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   8
         ToolTipText     =   "Change type of files"
         Top             =   4440
         Width           =   615
      End
      Begin VB.CommandButton comObrisi 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         ToolTipText     =   "Delete selected files in List"
         Top             =   3360
         Width           =   615
      End
      Begin VB.CommandButton comPrebaciSve 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         ToolTipText     =   "Move all files in List"
         Top             =   2880
         Width           =   615
      End
      Begin VB.CommandButton comPrebaci 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         ToolTipText     =   "Move selected files in List"
         Top             =   2400
         Width           =   615
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   120
         MultiSelect     =   2  'Extended
         TabIndex        =   4
         Top             =   2400
         Width           =   2775
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2775
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "TIFF 2 SIS configuration file (TIFF2SDI.CFG)"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   19
         Top             =   720
         Width           =   8655
         Begin VB.CommandButton comSaveConFile 
            Caption         =   "Save .CFG file"
            Height          =   255
            Left            =   6960
            TabIndex        =   29
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Z Gloabal Origin"
            Height          =   195
            Left            =   5280
            TabIndex        =   28
            Top             =   1200
            Width           =   1140
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Y Global Origin"
            Height          =   195
            Left            =   3600
            TabIndex        =   27
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "X Global Origin"
            Height          =   195
            Left            =   1920
            TabIndex        =   26
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Uor per Master"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Master Unit"
            Height          =   195
            Left            =   6960
            TabIndex        =   24
            Top             =   480
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Y Resolution"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   5280
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "X Resolution"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3600
            TabIndex        =   22
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Y Origin"
            Height          =   195
            Left            =   1920
            TabIndex        =   21
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "X Origin"
            Height          =   195
            Left            =   240
            TabIndex        =   20
            Top             =   480
            Width           =   555
         End
      End
      Begin VB.Label lblItemsCount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   50
         Top             =   4560
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "... items count"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   49
         Top             =   4560
         Width           =   1335
         WordWrap        =   -1  'True
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuNone 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit1 
      Caption         =   "Edit1"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuDeleteAll 
         Caption         =   "Delete All"
      End
   End
   Begin VB.Menu mnuEdit2 
      Caption         =   "Edit2"
      Visible         =   0   'False
      Begin VB.Menu mnuPrebaci 
         Caption         =   "Move"
      End
      Begin VB.Menu mnuPrebaciSve 
         Caption         =   "Move All"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewProcess 
         Caption         =   "View Process"
         Checked         =   -1  'True
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuHel 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmGlavna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const LB_SETHORIZONTALEXTENT = &H194
Private Sub comCancel_Click()
    Call Ending
End Sub

Private Sub ChkTIFF_Click()
    If ChkTIFF.Value = 0 Then
        If MsgBox("TIFF files will be deleted!" + Chr(13) + "Are you sure?", vbYesNo, "Alert") = vbNo Then
            ChkTIFF.Value = 1
        End If
    End If
End Sub
Private Sub ChkSIS_Click()
    If chkSIS.Value = 0 Then
        If MsgBox("SIS files will be deleted!" + Chr(13) + "Are you sure?", vbYesNo, "Alert") = vbNo Then
            chkSIS.Value = 1
        End If
    End If
End Sub
Private Sub ChkSJS_Click()
    If chkSJS.Value = 0 Then
        If MsgBox("SJS files will be deleted!" + Chr(13) + "Are you sure?", vbYesNo, "Alert") = vbNo Then
            chkSJS.Value = 1
        End If
    End If
End Sub

Private Sub comDefault95_Click()
    txtQuality.Text = "95"
End Sub

Private Sub ComFilter_Click()
    File1.Pattern = InputBox("Please, change type of files..." + Chr(13) + File1.Pattern)
End Sub

Private Sub comLOGFile_Click()
On Error GoTo ErrorHandler
    frmTextBox.rtbLogFile.LoadFile (App.Path + "\TIFF2SIS2SJS.LOG")
    frmTextBox.Show
    Exit Sub
ErrorHandler:
    a = MsgBox("LOG file doesn't exist!", vbExclamation, "Error")
End Sub

Private Sub comObrisi_Click()
Dim X As Long
Dim Y As Long
    Do While lstUlazneDatoteke.SelCount <> 0
        If lstUlazneDatoteke.Selected(0) Then
            lstUlazneDatoteke.RemoveItem (0)
        Else
            If lstUlazneDatoteke.Selected(i) Then
                lstUlazneDatoteke.RemoveItem (i)
                i = 0
            End If
        End If
        i = i + 1
    Loop
    X = Me.TextWidth(lstUlazneDatoteke.List(0) & " ") / Screen.TwipsPerPixelX
    For i = 0 To lstUlazneDatoteke.ListCount - 1
        Y = Me.TextWidth(lstUlazneDatoteke.List(i) & " ") / Screen.TwipsPerPixelX
        If X < Y Then
            X = Y
        End If
        SendMessageLong lstUlazneDatoteke.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    Next i
        lblItemsCount.Caption = Str(lstUlazneDatoteke.ListCount)
End Sub

Private Sub comObrisiSve_Click()
    lstUlazneDatoteke.Clear
    lblItemsCount.Caption = Str(lstUlazneDatoteke.ListCount)
End Sub

Private Sub comPrebaci_Click()
Dim X As Long
Dim Y As Long
    For i = 0 To File1.ListCount - 1
        If File1.Selected(i) Then
            File1.Selected(i) = False
            If Right$(Dir1.Path, 1) <> "\" Then
                lstUlazneDatoteke.AddItem (Dir1.Path + "\" + File1.List(i))
            Else
                lstUlazneDatoteke.AddItem (Dir1.Path + File1.List(i))
            End If
        End If
    Next i
    
    X = Me.TextWidth(lstUlazneDatoteke.List(0) & " ") / Screen.TwipsPerPixelX
    For i = 0 To lstUlazneDatoteke.ListCount - 1
        Y = Me.TextWidth(lstUlazneDatoteke.List(i) & " ") / Screen.TwipsPerPixelX
        If X < Y Then
            X = Y
        End If
        SendMessageLong lstUlazneDatoteke.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    Next i
        lblItemsCount.Caption = Str(lstUlazneDatoteke.ListCount)
End Sub

Private Sub comPrebaciSve_Click()
Dim X As Long
Dim Y As Long
    For i = 0 To File1.ListCount - 1
        If Right$(Dir1.Path, 1) <> "\" Then
            lstUlazneDatoteke.AddItem (Dir1.Path + "\" + File1.List(i))
        Else
            lstUlazneDatoteke.AddItem (Dir1.Path + File1.List(i))
        End If
    Next i
    
    X = Me.TextWidth(lstUlazneDatoteke.List(0) & " ") / Screen.TwipsPerPixelX
    For i = 0 To lstUlazneDatoteke.ListCount - 1
        Y = Me.TextWidth(lstUlazneDatoteke.List(i) & " ") / Screen.TwipsPerPixelX
        If X < Y Then
            X = Y
        End If
        SendMessageLong lstUlazneDatoteke.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    Next i
        lblItemsCount.Caption = Str(lstUlazneDatoteke.ListCount)
End Sub

Private Sub comSaveConFile_Click()
    modVanjskeDatoteke.Spremi_TIFF2SDI_Datoteku (True)
End Sub

Private Sub comStart_Click()
    If comStart.Caption = "Start" Then
        Call Pokreni
    Else
        comStart.Caption = "Please Wait... This may take a few minutes..."
        comStart.Enabled = False
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEdit2
    End If
End Sub

Private Sub File1_PatternChange()
    ComFilter.Caption = File1.Pattern
End Sub

Private Sub Form_Load()
Dim X As Long
Dim Y As Long

    modVanjskeDatoteke.Postavi_Defaultne_Vrijednosti
    modVanjskeDatoteke.Otvori_TIFF2SDI_Datoteku (App.Path + "\TIFF2SDI.CFG")
    modVanjskeDatoteke.Otvori_INI_Datoteku (App.Path + "\TIFF2SDI.INI")

    X = Me.TextWidth(lstUlazneDatoteke.List(0) & " ") / Screen.TwipsPerPixelX
    For i = 0 To lstUlazneDatoteke.ListCount - 1
        Y = Me.TextWidth(lstUlazneDatoteke.List(i) & " ") / Screen.TwipsPerPixelX
        If X < Y Then
            X = Y
        End If
        SendMessageLong lstUlazneDatoteke.hwnd, LB_SETHORIZONTALEXTENT, X, 0
    Next i
        lblItemsCount.Caption = Str(lstUlazneDatoteke.ListCount)
End Sub

Private Sub Form_Terminate()
    Close All
    End
End Sub

Public Sub Form_Unload(Cancel As Integer)
    Cancel = True
End Sub


Private Sub lstUlazneDatoteke_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call comObrisi_Click
    End If
End Sub

Private Sub lstUlazneDatoteke_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuEdit1
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
    Call Ending
End Sub

Private Sub mnuOpen_Click()
    CommonDialog1.DialogTitle = "Open .INI File"
    
    CommonDialog1.ShowOpen
    If CommonDialog1.FileName <> "" Then
        Call modVanjskeDatoteke.Otvori_INI_Datoteku(CommonDialog1.FileName)
    End If
End Sub

Private Sub mnuSave_Click()
    Call modVanjskeDatoteke.Spremi_INI_Datoteku(App.Path + "\TIFF2SDI.INI", True)
End Sub

Private Sub mnuSaveAs_Click()
    CommonDialog1.DialogTitle = "Save As File"
    CommonDialog1.ShowSave
    If CommonDialog1.FileName <> "" Then
        Call modVanjskeDatoteke.Spremi_INI_Datoteku(CommonDialog1.FileName, True)
    End If
End Sub

Private Sub mnuViewProcess_Click()
    mnuViewProcess.Checked = Not mnuViewProcess.Checked
End Sub

Private Sub Ending()
    Close All
    End
End Sub
Private Sub mnuPrebaci_Click()
    Call comPrebaci_Click
End Sub

Private Sub mnuPrebaciSve_Click()
    Call comPrebaciSve_Click
End Sub
Private Sub mnuDelete_Click()
    Call comObrisi_Click
End Sub

Private Sub mnuDeleteAll_Click()
    Call comObrisiSve_Click
End Sub

