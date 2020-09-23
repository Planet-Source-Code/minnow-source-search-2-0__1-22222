VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   3120
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   4560
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      Caption         =   "Ignore comments when searching"
      Height          =   372
      Left            =   360
      TabIndex        =   7
      ToolTipText     =   "Any commented lines in the project will be ignored."
      Top             =   600
      Value           =   1  'Checked
      Width           =   3492
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   372
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   372
      Left            =   3480
      TabIndex        =   1
      Top             =   2640
      Width           =   972
   End
   Begin VB.Frame Frame1 
      Height          =   2532
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4332
      Begin VB.CheckBox Check3 
         Caption         =   "Remove any unzipped files from the temporary unzip directory at exit (recommended)"
         Height          =   372
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Any commented lines in the project will be ignored."
         Top             =   1080
         Value           =   1  'Checked
         Width           =   3972
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Search on the Fly"
         Height          =   252
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "This allows code to be scanned as you drag and drop it."
         Top             =   240
         Value           =   1  'Checked
         Width           =   2412
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Save the Search Dictionary On Exit"
         Height          =   252
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Value           =   -1  'True
         Width           =   3732
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Let Me Save the Search Dictionary"
         Height          =   252
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   3492
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Save the Search Dictionary After Each Change"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   3972
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tIndex As Integer

'Triggered on checkbox click
Private Sub Check1_Click()

  'Enable the apply button
  cmdApply.Enabled = True

End Sub

Private Sub Check2_Click()

  'Enable the apply button
  cmdApply.Enabled = True

End Sub

'Triggered on exit click
Private Sub cmdExit_Click()

  'Unloads the form
  Unload Me

End Sub

'Triggered on apply button click
Private Sub cmdApply_Click()

  'Saves user's preferences
  SaveDictionary = tIndex
  If Check1.Value = vbChecked Then FlySearch = True Else FlySearch = False
  If Check2.Value = vbChecked Then IgnoreComments = True Else IgnoreComments = False
  If Check3.Value = vbChecked Then RemoveZips = True Else RemoveZips = False
  
  'Disables the apply button
  cmdApply.Enabled = False

End Sub

'Triggered on Form Load
Private Sub Form_Load()

  Dim x As Integer
  
  'Loads the defaut options
  For x = 0 To 2
    If Option1(x).Value = True Then tIndex = x
  Next x

End Sub

'Triggered on form unload
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim q As Variant
  
  'If either the control box or code has requested a shut down,
  'confirm it with the user.
  If UnloadMode <= 1 Then
    If cmdApply.Enabled Then
      q = MsgBox("Would you like to save your setting changes?", vbYesNo + vbQuestion)
      If q = vbYes Then Call cmdApply_Click
   End If
  End If

End Sub

'Triggered on option button click
Private Sub Option1_Click(Index As Integer)

  tIndex = Index
  cmdApply.Enabled = True

End Sub
