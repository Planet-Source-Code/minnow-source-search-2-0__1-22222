VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParentFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The collection of it's children
Private m_colChildFiles As Collection
'The path of the file
Dim m_strPath As String
'

Public Property Get Path() As String

  Path = m_strPath

End Property

Public Property Let Path(v_strPath As String)

  m_strPath = v_strPath

End Property

Public Property Get Count() As Long

  Count = m_colChildFiles.Count

End Property

Public Function Add(ByVal v_strPath As String) As Boolean
  
  Dim objChildFiles As New clsChildFile
  Dim x As Integer
  
  On Error Resume Next
  
  objChildFiles.Path = v_strPath
  
  If m_colChildFiles.Count > 0 Then
    For x = 1 To m_colChildFiles.Count
      If Item(x).Path = v_strPath Then Exit Function
    Next x
  End If
    
  m_colChildFiles.Add objChildFiles, v_strPath
  
  If Err.Number <> 0 Then
    Add = False
  Else
    Add = True
  End If
  
  Set objChildFiles = Nothing
 
End Function

Public Sub Clear()

  Dim x As Integer
  
  For x = m_colChildFiles.Count To 1 Step -1
    m_colChildFiles.Remove x
  Next x
  
End Sub

Public Function Remove(ByVal v_vntIndex As Variant)

  On Error Resume Next
  m_colChildFiles.Remove v_vntIndex
  
End Function

Public Function Item(ByVal v_vntIndex As Variant) As clsChildFile

  On Error Resume Next
  Set Item = m_colChildFiles.Item(v_vntIndex)
  If Err.Number <> 0 Then
    Set Item = Nothing
  End If
  
End Function

Private Sub Class_Initialize()

  Set m_colChildFiles = New Collection
  
End Sub





