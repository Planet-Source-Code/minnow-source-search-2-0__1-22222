VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Richsoft VBZip"
   ClientHeight    =   2016
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4800
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2016
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing © 2001"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "www.richsoftcomputing.btinternet.co.uk"
      DragIcon        =   "frmAbout.frx":030A
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Tag             =   "www.richsoftcomputing.btinternet.co.uk"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Image imgZip 
      Height          =   960
      Left            =   120
      Picture         =   "frmAbout.frx":045C
      Stretch         =   -1  'True
      Top             =   360
      Width           =   840
   End
   Begin VB.Label lblVBZip 
      BackStyle       =   0  'Transparent
      Caption         =   "VBZip"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblRichsoft 
      Caption         =   "Richsoft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'Richsoft Computing 2001
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'
'For latest information about this and other projects please visit my website:
'www.richsoftcomputing.btinternet.co.uk
'
'If you would like to make any comments/suggestions then please e-mail them to
'richsoftcomputing@btinternet.co.uk
'==============================================================================

'API Call which drives the Hyperlink
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub HyperJump(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub


Private Sub cmdOK_Click()
    'Close the About box
    Unload Me
End Sub

Private Sub lblWebsite_Click()
    'Jump to the website
    HyperJump lblWebsite.Tag
End Sub


Private Sub lblWebsite_DragDrop(Source As Control, X As Single, Y As Single)
    'If the mouse is over the label, the control
    'must be in drag mode. In this case, the
    'DragDrop event occurs when the mouse is
    'clicked.
    If Source Is lblWebsite Then
        With lblWebsite
            Call HyperJump(.Tag)
            .Font.Underline = False
            .ForeColor = vbButtonText
        End With
    End If
End Sub


Private Sub lblWebsite_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    'If the control is in drag mode, you can detect
    'MouseLeave easily by observing the State parameter.
    
    If State = vbLeave Then
        With lblWebsite
            .Drag vbEndDrag
            .FontUnderline = False
            .ForeColor = vbButtonText
        End With
    End If
End Sub


Private Sub lblWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Enter drag mode on the first MouseMove
    'allows easy detection of MouseLeave.
    
    With lblWebsite
        .ForeColor = vbHighlight
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub


