VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{36DBFD17-8A7D-4E36-A119-27B940B272CF}#1.0#0"; "VBZip_Control.ocx"
Begin VB.Form frmMain 
   Caption         =   "Source Search 2.0"
   ClientHeight    =   7284
   ClientLeft      =   132
   ClientTop       =   432
   ClientWidth     =   9936
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7284
   ScaleWidth      =   9936
   StartUpPosition =   2  'CenterScreen
   Begin VBZip_Control.RichsoftVBZip RZip 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1715
      _ExtentY        =   1715
   End
   Begin VB.Frame fraProjectList 
      Height          =   6372
      Left            =   240
      MouseIcon       =   "frmMain.frx":0E42
      OLEDropMode     =   1  'Manual
      TabIndex        =   7
      Top             =   480
      Width           =   9252
      Begin VB.PictureBox picProgressContainer 
         Height          =   252
         Left            =   3000
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   204
         ScaleWidth      =   5244
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4680
         Width           =   5292
         Begin VB.PictureBox picProgress 
            BackColor       =   &H00FFC0C0&
            BorderStyle     =   0  'None
            Height          =   372
            Left            =   -100
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   372
            ScaleWidth      =   372
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   372
         End
      End
      Begin MSComctlLib.ImageList PTIList 
         Left            =   480
         Top             =   3960
         _ExtentX        =   804
         _ExtentY        =   804
         BackColor       =   -2147483643
         ImageWidth      =   17
         ImageHeight     =   17
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   30
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":114C
               Key             =   "ClosedFolder"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16E6
               Key             =   "OpenFolder"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C80
               Key             =   "CLS"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DDA
               Key             =   "FRM"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F34
               Key             =   "CFRM"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":208E
               Key             =   "PFRM"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21E8
               Key             =   "BAS"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2342
               Key             =   "VBG"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":249C
               Key             =   "VBP"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25F6
               Key             =   "PAG"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2750
               Key             =   "CTL"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28AA
               Key             =   "CHECKCLS"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2E44
               Key             =   "CHECKFRM"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33DE
               Key             =   "CHECKCFRM"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3978
               Key             =   "CHECKPFRM"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F12
               Key             =   "CHECKBAS"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":44AC
               Key             =   "CHECKVBG"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4A46
               Key             =   "CHECKVBP"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4FE0
               Key             =   "CHECKPAG"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":557A
               Key             =   "CHECKCTL"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5B14
               Key             =   "CROSSCLS"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":60AE
               Key             =   "CROSSFRM"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6648
               Key             =   "CROSSCFRM"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6BE2
               Key             =   "CROSSPFRM"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":717C
               Key             =   "CROSSBAS"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7716
               Key             =   "CROSSVBG"
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7CB0
               Key             =   "CROSSVBP"
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":824A
               Key             =   "CROSSPAG"
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":87E4
               Key             =   "CROSSCTL"
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8D7E
               Key             =   "ZIP"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView PT 
         Height          =   4692
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2772
         _ExtentX        =   4890
         _ExtentY        =   8276
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin MSComctlLib.ListView SL 
         Height          =   4092
         Left            =   3000
         TabIndex        =   9
         Top             =   240
         Width           =   5292
         _ExtentX        =   9335
         _ExtentY        =   7218
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         OLEDropMode     =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Violation Phrase"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Severity"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Violation Context"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Violation in Procedure"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Violation Description"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label lblSearchingStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Not Currently Scanning"
         Height          =   192
         Left            =   4800
         TabIndex        =   10
         Top             =   4440
         Width           =   1644
      End
      Begin VB.Image imgCustomsize 
         Height          =   10800
         Left            =   2880
         MousePointer    =   9  'Size W E
         Picture         =   "frmMain.frx":9318
         Top             =   240
         Width           =   96
      End
   End
   Begin MSComDlg.CommonDialog CommDiag 
      Left            =   8640
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   393216
   End
   Begin VB.Frame fraSearchList 
      Height          =   1812
      Left            =   3960
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4092
      Begin VB.Frame fraSLSingle 
         Height          =   2172
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   2760
         Width           =   8172
         Begin VB.ComboBox cmbSeverity 
            Height          =   288
            ItemData        =   "frmMain.frx":9A27
            Left            =   6360
            List            =   "frmMain.frx":9A34
            OLEDropMode     =   1  'Manual
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Tag             =   "Normal"
            Top             =   480
            Width           =   1692
         End
         Begin VB.TextBox txtDescription 
            Height          =   972
            Left            =   120
            MultiLine       =   -1  'True
            OLEDropMode     =   1  'Manual
            TabIndex        =   19
            Top             =   1080
            Width           =   7932
         End
         Begin VB.TextBox txtPhrase 
            Height          =   288
            Left            =   120
            OLEDropMode     =   1  'Manual
            TabIndex        =   17
            Top             =   480
            Width           =   6132
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   192
            Left            =   120
            TabIndex        =   5
            Top             =   840
            Width           =   816
         End
         Begin VB.Label lblSeverity 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Severity"
            Height          =   192
            Left            =   6360
            TabIndex        =   4
            Top             =   240
            Width           =   588
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Search Phrase"
            Height          =   192
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1068
         End
      End
      Begin VB.Frame fraSLList 
         Height          =   2652
         Left            =   120
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         Top             =   120
         Width           =   8172
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   372
            Left            =   4680
            OLEDropMode     =   1  'Manual
            TabIndex        =   14
            Top             =   2160
            Width           =   972
         End
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   372
            Left            =   7080
            OLEDropMode     =   1  'Manual
            TabIndex        =   16
            Top             =   2160
            Width           =   972
         End
         Begin VB.CommandButton cmdUpdate 
            Caption         =   "Update"
            Enabled         =   0   'False
            Height          =   372
            Left            =   5880
            OLEDropMode     =   1  'Manual
            TabIndex        =   15
            Top             =   2160
            Width           =   972
         End
         Begin MSComctlLib.ListView PL 
            Height          =   1812
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   7932
            _ExtentX        =   13991
            _ExtentY        =   3196
            View            =   3
            LabelEdit       =   1
            SortOrder       =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            OLEDropMode     =   1
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            OLEDropMode     =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Key             =   "PHRASE"
               Text            =   "Search Phrase"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Key             =   "SEVERITY"
               Text            =   "Severity"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Key             =   "DESCRIPTION"
               Text            =   "Problem Description"
               Object.Width           =   8819
            EndProperty
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6972
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9492
      _ExtentX        =   16743
      _ExtentY        =   12298
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Project Search"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Dictionary"
            ImageVarType    =   2
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSSD 
         Caption         =   "&Save Search Dictionary"
      End
      Begin VB.Menu h2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileISD 
         Caption         =   "&Import Search Dictionary"
      End
      Begin VB.Menu mnuFileESD 
         Caption         =   "&Export Search Dictionary"
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSOptions 
         Caption         =   "&Options..."
      End
      Begin VB.Menu h3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSNow 
         Caption         =   "&Search Now"
      End
      Begin VB.Menu mnuSClearList 
         Caption         =   "&Clear Search List"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutSourceSearch 
         Caption         =   "Source Search"
      End
      Begin VB.Menu mnuAboutVBZip 
         Caption         =   "VBZip"
      End
   End
   Begin VB.Menu mnuP 
      Caption         =   "PopUpMenus"
      Visible         =   0   'False
      Begin VB.Menu mnuPSearchTree 
         Caption         =   "Search Tree"
         Begin VB.Menu mnuPSVNI 
            Caption         =   "View Node Info"
         End
         Begin VB.Menu mnuH2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPSSNodeNChildren 
            Caption         =   "Search This Node && Children"
         End
         Begin VB.Menu mnuPSSNode 
            Caption         =   "Search This Node Only"
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents SearchDictionary As clsDictionary
Attribute SearchDictionary.VB_VarHelpID = -1

'The key of the phrase list item being updated
Dim strUpdatingKey As String
'The indicator for the Custom Size Image being moved
Dim imgCustomsizeMouseDown As Boolean
'

'Triggered on the add button being clicked
Private Sub cmdAdd_Click()

  'Adds a new entry to the dictionary
  SearchDictionary.Add GenerateKey(PL), txtPhrase.Text, cmbSeverity.ListIndex, txtDescription.Text
  
  'Resets the fields
  txtPhrase.Text = ""
  txtDescription.Text = ""
  cmbSeverity.ListIndex = 1
  
  'Indicate that the dictionary has changed and should be saved
  DictionaryChanged = True
  
  'If we are supposed to save everytime, then do so
  If SaveDictionary = 0 Then SaveSearchDictionary SearchDictionary

End Sub

'Triggered on the remove button being clicked
Private Sub cmdRemove_Click()

  'Used incase the entry does not exist
  On Error Resume Next
  
  Dim tMsg As String
  Dim q As Variant
  
  'Setup the message
  tMsg = "Are you sure that you want to remove" & vbCrLf & vbCrLf _
      & PL.SelectedItem.Text & vbCrLf & vbCrLf _
      & "from the Source Search list?"
         
  'If PL.SelectedItem does not exist then an error would have
  'been generated
  If Err.Number <> 0 Then Exit Sub
  
  'Displat the message
  q = MsgBox(tMsg, vbYesNo + vbQuestion)
    
  If q = vbYes Then
    
    'Remove the enrty
    SearchDictionary.Remove PL.SelectedItem.Key
    'Indicate that the dictionary has changed and should be saved
    DictionaryChanged = True
    'If we are supposed to save everytime, then do so
    If SaveDictionary = 0 Then SaveSearchDictionary SearchDictionary
  
  End If

End Sub

'Triggered on the Update button click
Private Sub cmdUpdate_Click()

  'Reset each property
  SearchDictionary.Item(strUpdatingKey).Phrase = txtPhrase.Text
  SearchDictionary.Item(strUpdatingKey).Severity = cmbSeverity.ListIndex
  SearchDictionary.Item(strUpdatingKey).Description = txtDescription.Text
  
  'Indicate the update
  SearchDictionary.Update strUpdatingKey
  
  'Indicate that the dictionary has changed and should be saved
  DictionaryChanged = True
  'If we are supposed to save everytime, then do so
  If SaveDictionary = 0 Then SaveSearchDictionary SearchDictionary
  
End Sub

'Triggered on Form Load
Private Sub Form_Load()

  'Initializes the Search Dictionary
  Set SearchDictionary = New clsDictionary
  
  'Initializes the common dialog box
  CommDiag.Flags = cdlOFNCreatePrompt + cdlOFNOverwritePrompt + cdlOFNHideReadOnly
  CommDiag.DefaultExt = ".sdd"
  CommDiag.Filter = "Search Dictionary Definition (*.sdd)|*.sdd|All (*.*)|*.*"
  
  Call SetupWizard(Me, SearchDictionary, SL)
  
  'Runs global initialization function
  InitializeProject
  
  imgCustomsizeMouseDown = False
  
  'Loads the Search Dicitionary from the registry
  LoadSearchDictionary SearchDictionary, PL
    
  RemoveZips = True
  
End Sub

'Triggered when an unload request occurs
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Dim q As Variant
  
  'If either the control box or code has requested a shut down,
  'confirm it with the user.
  If UnloadMode <= 1 Then
    
    q = MsgBox("Are you sure that you want to quit?", vbYesNo + vbQuestion)
   
    If q = vbNo Then Cancel = 1
    
  End If
  
  'Prompts the user to save the dictionary if it has been changed.
  If q <> vbNo And SaveDictionary = 1 And DictionaryChanged Then
    SaveSearchDictionary SearchDictionary
  ElseIf q <> vbNo And SaveDictionary = 2 And DictionaryChanged Then
    q = MsgBox("Would you like to save the Seach Dictionary?", vbYesNo + vbQuestion)
    
    If q = vbYes Then SaveSearchDictionary SearchDictionary
    
  End If

End Sub

'Triggered on From Resize
Private Sub Form_Resize()

  'Calls the Resize Function
  ResizefrmMain

End Sub

'Triggered on form unload
Private Sub Form_Unload(Cancel As Integer)

  Dim lFolder As Variant
  
  'Saves the user's preferences
  SaveSetting AppTitle, "Options", "FlySearch", FlySearch
  SaveSetting AppTitle, "Options", "SaveDictionary", SaveDictionary
  SaveSetting AppTitle, "Options", "IgnoreComments", IgnoreComments
  SaveSetting AppTitle, "Options", "RemoveZips", RemoveZips
  
  'If the user chooses to, remove the temp unzips
  If RemoveZips Then
    On Error Resume Next
    For Each lFolder In FSO.GetFolder(Replace(App.Path & "\", "\\", "\") & "TempUnzip\").SubFolders
      FSO.DeleteFolder lFolder, True
    Next lFolder
  End If
  
End Sub

'Triggered when the user presses the mouse button on the resize image
'(on the project search tab.)
Private Sub imgCustomsize_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  'Sets a global variable
  imgCustomsizeMouseDown = True
  
End Sub

'Triggered when the user move the mouse over the image
'(on the project search tab.)
Private Sub imgCustomsize_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  'If the mouse is being de-pressed then we need to resize
  'some form components
  If imgCustomsizeMouseDown Then
    imgCustomsize.Left = imgCustomsize.Left + x
    ResizefrmMain
  End If

End Sub

'Triggered when the user releases the mouse button on the resize image
'(on the project search tab.)
Private Sub imgCustomsize_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  'Resets a global variable
  imgCustomsizeMouseDown = False

End Sub

'Triggered when the about menu is clicked
Private Sub mnuAboutSourceSearch_Click()

  frmAbout.Show vbModal, Me

End Sub

'Triggered when the about menu is clicked
Private Sub mnuAboutVBZip_Click()

  RZip.About

End Sub

'Triggered when the File, Export menu is clicked
Private Sub mnuFileESD_Click()

  'Clears the dialog filename
  CommDiag.FileName = ""
  CommDiag.ShowSave
  'If the dialog filename = "" then the user clicked cancel
  If CommDiag.FileName = "" Then Exit Sub
  
  'Calls the global export function
  ExportDictionary CommDiag.FileName, SearchDictionary

End Sub

'Triggered on File, Exit menu click
Private Sub mnuFileExit_Click()

  'Unloads the form
  Unload Me
  
End Sub

'Triggered on File, Import click
Private Sub mnuFileISD_Click()

  'Clears the dialog filename
  CommDiag.FileName = ""
  CommDiag.ShowOpen
  'If the dialog filename = "" then the user clicked cancel
  If CommDiag.FileName = "" Then Exit Sub
    
  'Calls the global import function
  ImportDictionary CommDiag.FileName, SearchDictionary, SL

End Sub

'Triggered on File, Save Dictionary menu click
Private Sub mnuFileSSD_Click()

  'Calls the global save dictionary function
  SaveSearchDictionary SearchDictionary

End Sub

'Triggered when the Popup menu is activated
Private Sub mnuPSearchTree_Click()

  'If the prject tree does not contain any project files, disable everything
  If blnProjectTreeInitialized = False Or PT.Nodes.Count < 1 Then
    mnuPSSNode.Enabled = False
    mnuPSSNodeNChildren.Enabled = False
    mnuPSVNI.Enabled = False
  Else
    mnuPSSNode.Enabled = True
    mnuPSVNI.Enabled = True
    
    'If the selected node does not have any children disable this menu item
    If PT.SelectedItem.Children = 0 Then mnuPSSNodeNChildren.Enabled = False Else mnuPSSNodeNChildren.Enabled = True
    
  End If

End Sub

'Triggered when the user clicks the search node popup menu
Private Sub mnuPSSNode_Click()

  'Searches a single file from the project tree
  SearchFile picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PT.SelectedItem

End Sub

'Triggered when the user clicks the search node and children popup menu
Private Sub mnuPSSNodeNChildren_Click()

  'Searches a single file from the project tree
  SearchFile picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PT.SelectedItem
  'Searches each child file of the main file
  SearchKeysChildren PT.SelectedItem.Key, PT, SL, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary

End Sub

'Triggered when the user clicks the show node info popup menu
Private Sub mnuPSVNI_Click()

  'Calls the routine
  ShowNodeInfo Me, SL, PT

End Sub

'Triggered when the user chooses to clear the Search List
Private Sub mnuSClearList_Click()

  Dim x As Integer
  Dim tImage As String
  
  'Clear the list
  SL.ListItems.Clear
  
  'Reset the icons of all of the tree items
  If blnProjectTreeInitialized And PT.Nodes.Count > 0 Then
    For x = 1 To PT.Nodes.Count
      tImage = PT.Nodes(x).Image
      If Len(tImage) > 5 Then
        If Left$(tImage, 5) = "CHECK" Or Left$(tImage, 5) = "CROSS" Then
          PT.Nodes(x).Image = Right$(tImage, Len(tImage) - 5)
        End If
      End If
    Next x
  End If

End Sub

'Triggered on Search menu click
Private Sub mnuSearch_Click()

  'Disables the Search Now menu if the appropriate elements are not set
  If PT.Nodes.Count = 0 Or Not blnProjectTreeInitialized Then mnuSNow.Enabled = False Else mnuSNow.Enabled = True
  If SL.ListItems.Count < 1 Then mnuSClearList.Enabled = False Else mnuSClearList.Enabled = True
    
End Sub

'Triggered on Search, Search Now menu click
Private Sub mnuSNow_Click()

  Dim x As Integer
  
  'If there are no nodes, get out
  If PT.Nodes.Count = 0 Then Exit Sub
  
  'Search Each node
  For x = 1 To PT.Nodes.Count
    
    SearchFile picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PT.Nodes(x)
    
  Next x

End Sub

'Triggered on Search, Options menu Click
Private Sub mnuSOptions_Click()

  'Loads the form
  Load frmOptions
  'Sets the specific form items
  If FlySearch Then frmOptions.Check1.Value = vbChecked Else frmOptions.Check1.Value = vbUnchecked
  If IgnoreComments Then frmOptions.Check2.Value = vbChecked Else frmOptions.Check2.Value = vbUnchecked
  If RemoveZips Then frmOptions.Check3.Value = vbChecked Else frmOptions.Check3.Value = vbUnchecked
  frmOptions.Option1(SaveDictionary).Value = True
  
  frmOptions.cmdApply.Enabled = False
  'Displays the form modally
  frmOptions.Show vbModal, Me

End Sub

'Triggered on phrase list column click
Private Sub PL_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  'Calls the global listview sort routine
  SortListView PL, ColumnHeader

End Sub

'Triggered on phrase list double click
Private Sub PL_DblClick()
  
  'If the list is empty, get out
  If PL.ListItems.Count = 0 Then Exit Sub
  
  'Set the text boxes as needed with the data
  txtPhrase.Text = PL.SelectedItem.Text
  txtDescription.Text = PL.SelectedItem.SubItems(2)
  
  Select Case PL.SelectedItem.SubItems(1)
    Case "Low"
      cmbSeverity.ListIndex = 0
    Case "Normal"
      cmbSeverity.ListIndex = 1
    Case "High"
      cmbSeverity.ListIndex = 2
  End Select
  
  'Sets the global variable that determines which item is being updated
  strUpdatingKey = PL.SelectedItem.Key
  
  'Enables the update button
  cmdUpdate.Enabled = True
  
End Sub

'Triggered on phrase list item click
Private Sub PL_ItemClick(ByVal Item As MSComctlLib.ListItem)

  'If the new key is not the key being updated, disable the update button
  If Item.Key <> strUpdatingKey Then cmdUpdate.Enabled = False
  
End Sub

'Triggered when the user double clicks the tree
Private Sub PT_DblClick()

  'Cancels the effect of the expanded state changing on Double Click
  PT.SelectedItem.Expanded = Not PT.SelectedItem.Expanded
  'Calls the show node info routine
  ShowNodeInfo Me, SL, PT

End Sub

'Triggered when the user wants to show the popup menu for the project tree
Private Sub PT_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  'Get out if there is no relavent data
  If Not blnProjectTreeInitialized Or PT.Nodes.Count < 1 Then Exit Sub
    
  'If it was a right click, responsd
  If Button = 2 Then
    'Ensures that the node that we will be preforming actions on is highlighted
    PT.Nodes(PT.SelectedItem.Key).Selected = True
    'Show's the popup menu
    PopupMenu mnuPSearchTree, , , , mnuPSVNI
  End If

End Sub

'Trigger on a new item being added to the search dictionary
Private Sub SearchDictionary_ItemAdded(NewItem As clsDefinition)

  'Add the data to the phrase list
  PL.ListItems.Add 1, NewItem.Key, NewItem.Phrase
  PL.ListItems(NewItem.Key).ListSubItems.Add(1, , cmbSeverity.List(NewItem.Severity)).ForeColor = QBColor(SeverityColor(NewItem.Severity))
  PL.ListItems(NewItem.Key).ListSubItems.Add 2, , NewItem.Description
  
  'Enables the remove button
  cmdRemove.Enabled = True

End Sub

'Triggered when an item is called from the search dictionary
'and that item does not exist (by key or index)
'The key passed to the dictionary is returned.
Private Sub SearchDictionary_ItemDoesNotExist(vItemKey As Variant)

  On Error Resume Next
  
  'Attempt to remove the item from the phrase list
  PL.ListItems.Remove vItemKey
  
  'Disable the remove button if the list is empty
  If PL.ListItems.Count = 0 Then cmdRemove.Enabled = False

End Sub

'Triggered when an item is removed from the dictionary
'The extinct item's key is returned.
Private Sub SearchDictionary_ItemRemoved(ItemKey As String)

  On Error Resume Next
  
  'Remove the key from the phrase list
  PL.ListItems.Remove ItemKey
  
  'Disable the remove button if the list is empty
  If PL.ListItems.Count = 0 Then cmdRemove.Enabled = False

End Sub

'Triggered when an item is updated in the dictionary
Private Sub SearchDictionary_ItemUpdated(ChangedItem As clsDefinition)

  'Updates the phrase list as required
  PL.ListItems(ChangedItem.Key).Text = ChangedItem.Phrase
  PL.ListItems(ChangedItem.Key).ListSubItems.Item(1).Text = cmbSeverity.List(ChangedItem.Severity)
  PL.ListItems(ChangedItem.Key).ListSubItems.Item(1).ForeColor = QBColor(SeverityColor(ChangedItem.Severity))
  PL.ListItems(ChangedItem.Key).ListSubItems.Item(2).Text = ChangedItem.Description
  
End Sub

'Triggered when the Search List column header is clicked
Private Sub SL_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

  'Calls the global list sorting
  SortListView SL, ColumnHeader

End Sub

'***********************************************************
'
'       OLE Drag and Drop Events
'
'   - OLE Drag - when a file(s) is/are being dragged
'       across the form, we set the project search tab
'       active make the form active.
'
'   - OLE Drop - when a file(s) is/are being dropped
'       onto the form, we accept them and take action.
'
'***********************************************************

Private Sub SL_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip

End Sub

Private Sub SL_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub TabStrip1_BeforeClick(Cancel As Integer)

  SetTabFrame TabStrip1.SelectedItem.Index, False

End Sub

Private Sub TabStrip1_Click()

  SetTabFrame TabStrip1.SelectedItem.Index, True
         
End Sub

Private Sub TabStrip1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip

End Sub

Private Sub TabStrip1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub txtDescription_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub txtPhrase_Change()

  If Trim(txtPhrase.Text) = "" Then cmdAdd.Enabled = False Else cmdAdd.Enabled = True
  
End Sub

Private Sub txtPhrase_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub cmbSeverity_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub cmdAdd_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub cmdRemove_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub cmdUpdate_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip

End Sub

Private Sub Form_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub fraProjectList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip

End Sub

Private Sub fraProjectList_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub fraSearchList_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub fraSLList_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub fraSLSingle_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub picProgressContainer_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip

End Sub

Private Sub picProgressContainer_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub PL_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub

Private Sub PT_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

  PublicOLEDragDrop Data, PT, picProgressContainer, picProgress, 99, lblSearchingStatus, SearchDictionary, SL, PTIList, RZip
   
End Sub

Private Sub PT_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

  PublicOLEDragOver Me

End Sub


