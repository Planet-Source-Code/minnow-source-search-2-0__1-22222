Attribute VB_Name = "ZipDll"
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

'Declarations of the library functions
Public Declare Function AddFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal Filename As String, ByVal StoreDirInfo As Boolean, ByVal DOS83 As Boolean, ByVal Action As Integer, ByVal CompressionLevel As Integer) As Boolean
Public Declare Function ExtractFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal Filename As String, ByVal ExtrDir As String, ByVal UseDirInfo As Boolean, ByVal Overwrite As Boolean, ByVal Action As Integer) As Boolean
Public Declare Function DeleteFile Lib "zipit.dll" (ByVal ZipFilename As String, ByVal Filename As String) As Boolean

