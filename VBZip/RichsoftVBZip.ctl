VERSION 5.00
Begin VB.UserControl RichsoftVBZip 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   972
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   972
   HasDC           =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   972
   ScaleWidth      =   972
   ToolboxBitmap   =   "RichsoftVBZip.ctx":0000
   Begin VB.Frame fraBorder 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin VB.Image imgZip 
         Height          =   384
         Left            =   120
         Picture         =   "RichsoftVBZip.ctx":0312
         Top             =   240
         Width           =   384
      End
   End
End
Attribute VB_Name = "RichsoftVBZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'Set up the private atrributes
Private ZipFilename As String
Private CompLevel As ZipLevel
Private DOS83Format As Boolean
Private Recurse As Boolean

'Set up the file collection
Private Archive As Collection

'Events
Event OnArchiveUpdate()
Event OnZipProgress(ByVal Percentage As Integer, ByVal Filename As String)
Event OnZipComplete(ByVal Successful As Long)
Event OnUnzipProgress(ByVal Percentage As Integer, ByVal Filename As String)
Event OnUnzipComplete(ByVal Successful As Long)
Event OnDeleteProgress(ByVal Percentage As Integer, ByVal Filename As String)
Event OnDeleteComplete(ByVal Successful As Long)

'Actions
Public Enum ZipAction
    zipDefault = 1
    zipFreshen = 2
    zipUpdate = 3
End Enum

'Compression Level
Public Enum ZipLevel
    zipStore = 0
    zipLevel1 = 1
    zipSuperFast = 2
    zipFast = 3
    zipLevel4 = 4
    zipNormal = 5
    zipLevel6 = 6
    zipLevel7 = 7
    zipLevel8 = 8
    zipMax = 9
End Enum

Public Function ConvertBytesToString(Bytes As Long) As String
    'Turns a number representing the number of bytes
    'into a string, bytes, KB, MB
    
    Select Case (Bytes / 1024)
        Case Is < 0.2
            ConvertBytesToString = Format(Bytes, "###,##0") & " bytes"
        
        Case Is < 512
            ConvertBytesToString = Format(Bytes / 1024, "###,##0.0") + "KB"
            
        Case Else
            ConvertBytesToString = Format(Bytes / (1024 ^ 2), "###,##0.0") + "MB"
            
    End Select
        
End Function
Private Function FindFiles(Files As Collection, ByVal Recurse As Boolean)
    'Finds all the files matching the specification
    
    '*******************************************
    'RECURSIVE FOLDER SEARCH NOT YET IMPLEMENTED
    '*******************************************
    Dim Result As New Collection
    Dim i As Long
    For i = 1 To Files.Count
        'Debug.Print Files(i)
        'Parse the file specification to find the path
        Path = ParsePath(Files(i))
        'Find the files matching the pattern
        r = Dir$(Files(i), Attributes)
        Do Until r = ""
            'Add the file to the new file list collection
            Result.Add Path & r
            'Move on to next file, if one exists
            r = Dir$()
        Loop
    Next i
    
    Set FindFiles = Result
    
End Function

Public Function Add(Files As Collection, ByVal Action As ZipAction, ByVal StorePathInfo As Boolean, ByVal RecurseSubFolders As Boolean, ByVal UseDOS83 As Boolean, ByVal CompressionLevel As ZipLevel) As Long
    'Adds the specified files to the archive
    Dim ArchiveFilename As String
    ArchiveFilename = ZipFilename
    Dim i As Long
    Dim Result As Long
    Dim FilesToAdd As Collection
    
    'Check to see if there are any files in the archive
    'if not delete the file so there are not error messages
    
    If GetEntryNum = 0 Then
        If Dir$(ArchiveFilename) <> "" Then
            Kill ArchiveFilename
        End If
    End If
    
    'Find all the files to add, recursing folders if selected
    Set FilesToAdd = FindFiles(Files, RecurseSubFolders)
    
    'Loop through the files adding them to the archive
    For i = 1 To FilesToAdd.Count
        'Debug.Print "Trying to Add " & FilesToAdd(i)
        RaiseEvent OnZipProgress((100 * (i / (FilesToAdd.Count))), ParseFilename(FilesToAdd(i)))
        DoEvents
        If AddFile(ArchiveFilename, FilesToAdd(i), StorePathInfo, UseDOS83, Action, CompressionLevel) Then
            'File successfully extracted
            Result = Result + 1
        Else
            'File did not extract for some reason
        End If
    Next i
    
    RaiseEvent OnZipComplete(Result)
    'If any file was added update the archive
    If Result > 0 Then
        Read
        RaiseEvent OnArchiveUpdate
    End If
End Function


Public Function Delete(Files As Collection) As Long
    'Deletes the files specified in the collection
    'Returns the number of files deleted
    Dim FilesToDelete As Collection
    Dim ArchiveFilename As String
    ArchiveFilename = ZipFilename
    Dim i As Long
    Dim Result As Long
    
    'First find the files which match the patterns
    'specified in the collection
    Set FilesToDelete = SelectFiles(Files)
    
    'Extract each file in turn
    For i = 1 To FilesToDelete.Count
        'Debug.Print "Trying to Delete " & FilesToDelete(i)
        RaiseEvent OnDeleteProgress((100 * (i / (FilesToDelete.Count))), ParseFilename(FilesToDelete(i)))
        DoEvents
        'Check to see if we are deleting the last entry
        'if so just delete the archive
        If (GetEntryNum - Result) > 1 Then
            If DeleteFile(ArchiveFilename, FilesToDelete(i)) Then
                'File successfully extracted
                Result = Result + 1
            Else
                'File did not extract for some reason
            End If
        Else
            Kill ArchiveFilename
            Result = Result + 1
        End If
    Next i
    
    RaiseEvent OnDeleteComplete(Result)
    'If any file was deleted update the archive
    If Result > 0 Then
        Read
        RaiseEvent OnArchiveUpdate
    End If
    
    Delete = Result
    
End Function

Public Function Extract(Files As Collection, ByVal Action As ZipAction, ByVal UsePathInfo As Boolean, ByVal Overwrite As Boolean, ByVal Path As String) As Long
    'Extracts the files specified in the collection
    'Returns the number of files extracted
    Dim FilesToExtract As Collection
    Dim ArchiveFilename As String
    ArchiveFilename = ZipFilename
    Dim i As Long
    Dim Result As Long
    
    'MsgBox Files.Count
    
    'First find the files which match the patterns
    'specified in the collection
    Set FilesToExtract = SelectFiles(Files)
    
    'Check to if there is anything to do
    'if there is create the output path if it does not exist
    
    '************
    'TO IMPLEMENT
    '************
    
    'Extract each file in turn
    'MsgBox FilesToExtract.Count
    
    For i = 1 To FilesToExtract.Count
        'Debug.Print "Trying to Extract " & FilesToExtract(i) & " to " & Path
        RaiseEvent OnUnzipProgress((100 * (i / (FilesToExtract.Count))), ParseFilename(FilesToExtract(i)))
        DoEvents
        If ExtractFile(ArchiveFilename, CStr(FilesToExtract(i)), Path, UsePathInfo, Overwrite, Action) Then
            'File successfully extracted
            Result = Result + 1
            'Debug.Print "Successful Extract " & FilesToExtract(i) & " to " & Path
        Else
            'Debug.Print "NOT Successful Extract " & FilesToExtract(i) & " to " & Path
        End If
    Next i
    
    RaiseEvent OnUnzipComplete(Result)
    Extract = Result
    
End Function

Public Property Get Filename() As String
    Filename = ZipFilename
End Property

Public Property Let Filename(ByVal New_Filename As String)
    Dim r As Long
    Dim i As Long
    'Called when the filename is updated
    ZipFilename = New_Filename
    PropertyChanged "Filename"
    'Read in the contents of the file
    r = Read
    'Raise the update event
    RaiseEvent OnArchiveUpdate
End Property

Public Function GetEntry(ByVal Index As Long) As ZipFileEntry
    Set GetEntry = Archive(Index)
End Function
Public Function GetEntryNum() As Long
    GetEntryNum = Archive.Count
End Function

Private Function SelectFiles(Files As Collection) As Collection
    'Selects files from a wildcard specification
    'Wildcards only corrispond to the filename and not the path
    Dim i As Long
    Dim j As Long
    Dim Result As New Collection
    'Loop through the collection looking at each entry
    For i = 1 To Files.Count
        'Loop through the files in the archive checking the pattern
        'Debug.Print Files(i)
        For j = 1 To GetEntryNum()
            'Check the pattern, ignoring case
            'Debug.Print "     " & GetEntry(j).Filename & "  -  " & ParseFilename(GetEntry(j).Filename)
            If LCase$(GetEntry(j).Filename) = LCase$(Files(i)) Then
                'Its a match so add it to the new collection
                Result.Add GetEntry(j).Filename
            End If
        Next j
    Next i
    Set SelectFiles = Result
End Function




























Public Function ParsePath(Path As String) As String
    'Takes a full file specification and returns the path
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            'Add the correct path separator for the input
            If Mid$(Path, A, 1) = "\" Then
                ParsePath = LCase$(Left$(Path, A - 1) & "\")
            Else
                ParsePath = LCase$(Left$(Path, A - 1) & "/")
            End If
            Exit Function
        End If
    Next A
End Function

Public Function ParseFilename(ByVal Path As String) As String
    'Takes a full file specification and returns the path
    For A = Len(Path) To 1 Step -1
        If Mid$(Path, A, 1) = "\" Or Mid$(Path, A, 1) = "/" Then
            ParseFilename = Mid$(Path, A + 1)
            Exit Function
        End If
    Next A
    ParseFilename = Path
End Function

Private Sub UserControl_Initialize()
    'Create a new Archive Collection
    Set Archive = New Collection
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Get properties out of storage
    ZipFilename = PropBag.ReadProperty("Filename", "")
End Sub


Private Sub UserControl_Resize()
    UserControl.Size 975, 975
End Sub

Public Sub About()
    'Show the about box
    frmAbout.Show 1
End Sub

Public Function Read() As Long
    'Reads the archive and places each file into a collection
    Dim Sig As Long
    Dim ZipStream As Integer
    Dim Res As Long
    Dim zFile As ZipFile
    Dim Name As String
    Dim i As Integer
    
    'If the filename is empty return a empty file list
    If ZipFilename = "" Then
        Read = 0
        'Remove any files still in the list
        For i = Archive.Count To 1 Step -1
            Archive.Remove i
        Next i
        Exit Function
    End If
    
    'Clears the collection
    'begin
    'Archive.Clear;
    For i = Archive.Count To 1 Step -1
        Archive.Remove i
    Next i
    
    'Opens the archive for binary access
    ZipStream = FreeFile
    Open ZipFilename For Binary As ZipStream
    'Loop through archive
    Do While True
        Get ZipStream, , Sig
        'See if the file header has been found
              If Sig = LocalFileHeaderSig Then
                    'Read each part of the file header
                    Get ZipStream, , zFile.Version
                    Get ZipStream, , zFile.Flag
                    Get ZipStream, , zFile.CompressionMethod
                    Get ZipStream, , zFile.Time
                    Get ZipStream, , zFile.Date
                    Get ZipStream, , zFile.CRC32
                    Get ZipStream, , zFile.CompressedSize
                    Get ZipStream, , zFile.UncompressedSize
                    Get ZipStream, , zFile.FileNameLength
                    Get ZipStream, , zFile.ExtraFieldLength
                    'Get the filename
                    'Set up a empty string so the right number of
                    'bytes is read
                    Name = String$(zFile.FileNameLength, " ")
                    Get ZipStream, , Name
                    zFile.Filename = Mid$(Name, 1, zFile.FileNameLength)
                    'Move on through the archive
                    'Skipping extra space, and compressed data
                    Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                    'Add the fileinfo to the collection
                    AddEntry zFile
              Else
              'Debug.Print Sig
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    'All the filenames have been found so
                    'exit the loop
                    Exit Do
                'End
                Else
                If Sig = EndCentralDirSig Then
                    'Exit the loop
                    Exit Do
                End If
                End If
            End If
        Loop
        'Close the archive
        Close ZipStream
        'Return the number of files in the archive
        Read = Archive.Count

    'Fire the update event
    RaiseEvent OnArchiveUpdate
End Function

Private Sub AddEntry(zFile As ZipFile)
    Dim xFile As New ZipFileEntry
    'Adds a file from the archive into the collection
    '**It does not add entry that are just folders**
    If ParseFilename(zFile.Filename) <> "" Then
        xFile.Version = zFile.Version
        xFile.Flag = zFile.Flag
        xFile.CompressionMethod = zFile.CompressionMethod
        xFile.CRC32 = zFile.CRC32
        xFile.FileDateTime = GetDateTime(zFile.Date, zFile.Time)
        xFile.CompressedSize = zFile.CompressedSize
        xFile.UncompressedSize = zFile.UncompressedSize
        xFile.FileNameLength = zFile.FileNameLength
        xFile.Filename = zFile.Filename
        xFile.ExtraFieldLength = zFile.ExtraFieldLength
    End If
    Archive.Add xFile
End Sub

Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string
    
    Dim r As Long
    Dim FTime As FileTime
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String
    
    'Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Convert the file time into a standard time
    r = FileTimeToSystemTime(FTime, Sys)

    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond

    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function
Private Sub UserControl_Terminate()
    'Clean up the Archive
    Set Archive = Nothing
End Sub


Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Put properties into storage
    PropBag.WriteProperty "Filename", ZipFilename, ""
End Sub

