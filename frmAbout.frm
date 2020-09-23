VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   6072
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   7212
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6072
   ScaleWidth      =   7212
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3612
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   6371
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAbout.frx":0E42
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6972
      Begin VB.Image Image1 
         Height          =   1680
         Left            =   20
         Picture         =   "frmAbout.frx":0ECD
         Top             =   120
         Width           =   6900
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   372
      Left            =   6120
      TabIndex        =   0
      Top             =   5640
      Width           =   972
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()

  Unload Me

End Sub

Private Sub Form_Load()

  RichTextBox1.TextRTF = ReleaseNotesInfo

End Sub
