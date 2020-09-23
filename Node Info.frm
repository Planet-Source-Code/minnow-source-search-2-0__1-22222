VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNodeInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Source Search Node Information Dialog"
   ClientHeight    =   5256
   ClientLeft      =   36
   ClientTop       =   336
   ClientWidth     =   8700
   Icon            =   "Node Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5256
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5172
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   8532
      Begin VB.ListBox lstParents 
         Height          =   1008
         ItemData        =   "Node Info.frx":0E42
         Left            =   4320
         List            =   "Node Info.frx":0E49
         TabIndex        =   7
         Top             =   3960
         Width           =   4092
      End
      Begin VB.ListBox lstChildren 
         Height          =   1008
         ItemData        =   "Node Info.frx":0E57
         Left            =   120
         List            =   "Node Info.frx":0E5E
         TabIndex        =   6
         Top             =   3960
         Width           =   4092
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Open This Folder"
         Height          =   372
         Left            =   6720
         TabIndex        =   5
         Top             =   1440
         Width           =   1692
      End
      Begin VB.TextBox txtPath 
         Height          =   288
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1080
         Width           =   4092
      End
      Begin VB.TextBox txtDisplayName 
         Height          =   288
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   4092
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Open This File"
         Height          =   372
         Left            =   4320
         TabIndex        =   2
         Top             =   1440
         Width           =   1692
      End
      Begin MSComctlLib.TreeView TV 
         Height          =   3372
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4092
         _ExtentX        =   7218
         _ExtentY        =   5948
         _Version        =   393217
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin MSComctlLib.ListView LV 
         Height          =   1452
         Left            =   4320
         TabIndex        =   8
         Top             =   2160
         Width           =   4092
         _ExtentX        =   7218
         _ExtentY        =   2561
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Violation Level"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Number of Violations"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Violations"
         Height          =   192
         Left            =   4320
         TabIndex        =   13
         Top             =   1920
         Width           =   708
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Node's Immediate Parents"
         Height          =   192
         Left            =   4320
         TabIndex        =   12
         Top             =   3720
         Width           =   1908
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Node's Immediate Children"
         Height          =   192
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   1944
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Node's File Path"
         Height          =   192
         Left            =   4320
         TabIndex        =   10
         Top             =   840
         Width           =   1188
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Node's Display Name"
         Height          =   192
         Left            =   4320
         TabIndex        =   9
         Top             =   240
         Width           =   1596
      End
   End
End
Attribute VB_Name = "frmNodeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

  'Opens the folder of the root node
  ShellExecute 0&, vbNullString, Mid$(FSO.GetFile(txtPath.Text).Path, 1, Len(FSO.GetFile(txtPath).Path) - Len(FSO.GetFile(txtPath).Name)), vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub Command2_Click()

  'Opens the root node in VB (or the default program)
  ShellExecute 0&, vbNullString, txtPath.Text, vbNullString, vbNullString, vbNormalFocus

End Sub

Private Sub Form_Load()

  'Sets default options
  LV.ListItems.Add(, "Low", "Low").ListSubItems.Add , , "0"
  LV.ListItems.Add(, "Normal", "Normal").ListSubItems.Add , , "0"
  LV.ListItems.Add(, "High", "High").ListSubItems.Add , , "0"
  LV.ListItems.Add(, "Total", "Total").ListSubItems.Add , , "0"
  
  TV.ImageList = frmMain.PTIList
    
End Sub

