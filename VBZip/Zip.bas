Attribute VB_Name = "ZipUtils"
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

'Zip file format type
Type ZipFile
  Version As Integer                    ': WORD;
  Flag As Integer                       ': WORD;
  CompressionMethod As Integer          ': WORD;
  Time As Integer                       ': WORD;
  Date As Integer                       ': WORD;
  CRC32 As Long                      ': Longint;
  CompressedSize As Long             ': Longint;
  UncompressedSize As Long           ': Longint;
  FileNameLength As Integer             ': WORD;
  ExtraFieldLength As Integer           ': WORD;
  Filename As String                 ': String;
End Type

'Zip file constants
Public Const LocalFileHeaderSig = &H4034B50
Public Const CentralFileHeaderSig = &H2014B50
Public Const EndCentralDirSig = &H6054B50

'App constants
Public Const APP_TITLE = "Richsoft Zipit 1.0"

'File dates/times functions and types
Public Declare Function DosDateTimeToFileTime Lib "kernel32" (ByVal wFatDate As Long, ByVal wFatTime As Long, lpFileTime As FileTime) As Long
Public Type FileTime
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FileTime, lpSystemTime As SYSTEMTIME) As Long
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type


'Declarations below this point are for future use...


Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Const MAX_PATH = 260
Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FileTime
        ftLastAccessTime As FileTime
        ftLastWriteTime As FileTime
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

