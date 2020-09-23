VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Definition Key (not carried over from session to session)
Dim m_strKey As String
'The definition description
Dim m_strDescription As String
'The Severity Rating
Dim m_intSeverity As Integer
'The Search Phrase
Dim m_strPhrase As String

Public Property Get Key() As String

  Key = m_strKey

End Property

Public Property Let Key(v_strKey As String)

  m_strKey = v_strKey

End Property

Public Property Get Description() As String

  Description = m_strDescription

End Property

Public Property Let Description(v_strDescription As String)

  m_strDescription = v_strDescription

End Property

Public Property Get Severity() As Integer

  Severity = m_intSeverity

End Property

Public Property Let Severity(v_intSeverity As Integer)

  m_intSeverity = v_intSeverity

End Property

Public Property Get Phrase() As String

  Phrase = m_strPhrase

End Property

Public Property Let Phrase(v_strPhrase As String)

  m_strPhrase = v_strPhrase

End Property


