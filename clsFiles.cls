VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The collection of all parents
Private m_colParentFiles As Collection

Public Function IsParent(v_strPath As String) As Boolean

  On Error GoTo Err_Handle
  
  Dim lItem As clsParentFile
  
  Set lItem = m_colParentFiles.Item(v_strPath)
  Set lItem = Nothing
  
  IsParent = True
  Exit Function
  
Err_Handle:
  Set lItem = Nothing
  IsParent = False
  
End Function

Public Property Get Count() As Long

  Count = m_colParentFiles.Count

End Property

Public Function Add(ByVal v_strPath As String) As Boolean
  
  Dim objParentFiles As New clsParentFile
  Dim x As Integer
  
  On Error Resume Next
  
  objParentFiles.Path = v_strPath
  
  If m_colParentFiles.Count > 0 Then
    For x = 1 To m_colParentFiles.Count
      If Item(x).Path = v_strPath Then Exit Function
    Next x
  End If
  
  m_colParentFiles.Add objParentFiles, v_strPath
  
  If Err.Number <> 0 Then
    Add = False
  Else
    Add = True
  End If
  
  Set objParentFiles = Nothing
 
End Function

Public Sub Clear()

  Dim x As Integer
  
  For x = m_colParentFiles.Count To 1 Step -1
    m_colParentFiles.Remove x
  Next x
  
End Sub

Public Function Remove(ByVal v_vntIndex As Variant)

  On Error Resume Next
  m_colParentFiles.Remove v_vntIndex
  
End Function

Public Function Item(ByVal v_vntIndex As Variant) As clsParentFile

  On Error Resume Next
  Set Item = m_colParentFiles.Item(v_vntIndex)
  If Err.Number <> 0 Then
    Set Item = Nothing
  End If
  
End Function

Private Sub Class_Initialize()

  Set m_colParentFiles = New Collection
  
End Sub




