VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmTextBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View LOG file"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmTextBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7875
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbLogFile 
      Height          =   4905
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   8652
      _Version        =   393217
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmTextBox.frx":6852
   End
End
Attribute VB_Name = "frmTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
