VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The collection of definitions
Private m_colDefinition As Collection

Public Event ItemAdded(NewItem As clsDefinition)
Public Event ItemRemoved(ItemKey As String)
Public Event ItemUpdated(ChangedItem As clsDefinition)
Public Event ItemDoesNotExist(vItemKey As Variant)

'Returns the number of definitions
Public Property Get Count() As Long

  Count = m_colDefinition.Count

End Property

'Method to trigger the update event
Public Function Update(ByVal v_vntIndex As Variant)

  Dim tItem As clsDefinition
  Dim tKey As String
  
  'Gets the key (if the index was submitted
  tKey = GetKey(v_vntIndex)
  
  'If the key does not exist then exit method
  If tKey = "" Then Exit Function
  
  'Setup the new item
  Set tItem = m_colDefinition.Item(tKey)
  
  'Raise the event
  RaiseEvent ItemUpdated(tItem)

End Function

'Method to add a new definition
Public Function Add(ByVal v_strKey As String, Optional ByVal v_strPhrase As String = "", Optional ByVal v_intSeverity As Integer = 1, Optional ByVal v_strDescription As String = "") As Boolean
  
  Dim objDefinition As New clsDefinition
  Dim x As Integer
  
  On Error Resume Next
  
  'Setup each property of the definistion
  objDefinition.Key = v_strKey
  objDefinition.Phrase = v_strPhrase
  objDefinition.Description = v_strDescription
  objDefinition.Severity = v_intSeverity
  
  'Add the definition
  m_colDefinition.Add objDefinition, v_strKey
  
  'If there has been an error, return false, otherwise true
  If Err.Number <> 0 Then
    Add = False
  Else
    Add = True
  End If
  
  'Raise the item added event
  RaiseEvent ItemAdded(objDefinition)
  
  Set objDefinition = Nothing
 
End Function

'Method to clear the definitions
Public Sub Clear()

  Dim x As Integer
  
  'Remove each definition
  For x = m_colDefinition.Count To 1 Step -1
    m_colDefinition.Remove x
  Next x
  
End Sub

'Method to remove a single definition
Public Function Remove(ByVal v_vntIndex As Variant)

  Dim tKey As String
  
  On Error Resume Next
  
  'Get the key of the definition if it's index was passed
  tKey = GetKey(v_vntIndex)
  
  'If the key does not exist then get out
  If tKey = "" Then Exit Function
  
  'Remove the definition
  m_colDefinition.Remove v_vntIndex
  
  'Raise the event
  RaiseEvent ItemRemoved(tKey)
  
End Function

'Method to return an item
Public Function Item(ByVal v_vntIndex As Variant) As clsDefinition

  On Error Resume Next
  'Get the item from the collection
  Set Item = m_colDefinition.Item(v_vntIndex)
  
  'If there was an error, handle it
  If Err.Number <> 0 Then
    RaiseEvent ItemDoesNotExist(v_vntIndex)
    Set Item = Nothing
  End If
  
End Function

'Verifies and returns a key to a definition
Private Function GetKey(vIndex As Variant) As String

  Dim tItem As clsDefinition
  
  'Attemp to set the item by the key
  Set tItem = m_colDefinition.Item(vIndex)
  
  'If there was an error, raise the event
  If Err.Number <> 0 Then
    RaiseEvent ItemDoesNotExist(vIndex)
    GetKey = ""
  Else
    'Rerturn the items key
    GetKey = tItem.Key
  End If

  Set tItem = Nothing

End Function

'Triggered on class initialize
Private Sub Class_Initialize()

  'Initializes the collection
  Set m_colDefinition = New Collection
  
End Sub





