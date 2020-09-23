Attribute VB_Name = "PublicFunctions"
Option Explicit

'The file scripting reference
Public FSO As New Scripting.FileSystemObject

'Used to open files by thier default program
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'The name of the regeistry section
Public Const AppTitle = "Source Search"

'Gets the hWnd of the current system window
Public Declare Function GetForegroundWindow Lib "user32" () As Long

'The class of parent files
Public ParentFiles As New clsFiles

'Contains the QBColor for each level of severity
Public SeverityColor(3) As Integer
'Contains the description for each level of severity
Public SeverityDescription(3) As String
'Contains the valid VB file extenions
Public ValidVBFiles(6) As String
'Contains the keys of all root nodes in the project tree
Public RootNodes() As String
'Contains the possible procedure names
Public ProcedureNames(2) As String

'The key to indicate if the project tree has been initialized
Public blnProjectTreeInitialized As Boolean
'vbQ = chr$(34) = "
Public vbQ As String
'Indicates the method (interval) that should be used to save the dictionary
Public SaveDictionary As Integer
'Indicates if the project should be search as it is dragged and dropped
Public FlySearch As Boolean
'The percent of the file that has been scanned
Public SearchPercent As Single
'The procedure that is being searched
Public CurrentProcedure As String
'Indicates that a procedure has been found
Public FoundProcedure As Boolean
'Indicates that a violation was found
Public ViolationFound As Boolean
'Indicates that the dictionary changed
Public DictionaryChanged As Boolean
'Indicates if the user would like to ignore commented lines while searching
Public IgnoreComments As Boolean
'Indiacates if the user would like to remove any temp unzips on exit
Public RemoveZips As Boolean

'A file detail type
Public Type typFile
    Ext As String
    Name As String
    Type As String
    vbName As String
    FilePath As String
    Size As Single
    FullName As String
End Type
'

'Creates a temporary, unique directory to unzip a file to
Private Function CreateTempUnzipDirectory(lFullFile As String) As String

  Dim ZipDir As String
  Dim lFile As typFile
  Dim lPath As String
  Dim x As Integer
   
  'Build the file type
  lFile = BuildFileType(lFullFile)
  'Sets the default directory to unzip to
  ZipDir = "TempUnzip"
  lPath = Replace(App.Path & "\", "\\", "\") & ZipDir & "\"
  
  'If the folder doesn't exsist, create it
  If Not FSO.FolderExists(lPath) Then FSO.CreateFolder lPath
  
  'If the file name as a subfolder already exists, create a unique one
  'i.e. - you have a zip file - abcde.zip
  'We check to see if ...\ZipDir\abcde\ exists
  'If it does, we try ...\ZipDir\abcde1\ until we find one that does not exist
  If FSO.FolderExists(lPath & lFile.Name & "\") Then
    x = 1
    Do While FSO.FolderExists(lPath & lFile.Name & x & "\")
      x = x + 1
    Loop
    CreateTempUnzipDirectory = lPath & lFile.Name & x & "\"
  Else
    'Returns the name of the new directory
    CreateTempUnzipDirectory = lPath & lFile.Name & "\"
  End If
  
  'Create the new folder
  FSO.CreateFolder CreateTempUnzipDirectory
    
End Function

'Used to Unzip a Zip File and add the contents to the tree
Private Function UnzipFile(TV As TreeView, lFullFile As String, tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary, LV As ListView, RZip As RichsoftVBZip) As String

  Dim lPath As String
  Dim lFiles As New Collection
  Dim x As Integer
  Dim lFile As Variant
        
  'Sets the initial state
  UnzipFile = ""
  'Creates a directory to unzip to
  lPath = CreateTempUnzipDirectory(lFullFile)
  
  'If the folder does not exist, then quit
  If Not FSO.FolderExists(lPath) Then Exit Function
  
  RZip.FileName = lFullFile
  
  'If the zip file is empty, quit
  If RZip.GetEntryNum = 0 Then
    RZip.FileName = ""
    Exit Function
  End If
      
  'Adds the name of each file to a file collection
  For x = 1 To RZip.GetEntryNum
    
    If Trim(RZip.GetEntry(x).FileName) <> "" Then
      lFiles.Add RZip.GetEntry(x).FileName
    End If
  
  Next x
  
  'Unzips the file collection
  RZip.Extract lFiles, zipDefault, True, True, lPath
  
  'If no files have been unzipped, get out
  If FSO.GetFolder(lPath).Files.Count = 0 Then
    RZip.FileName = ""
    Exit Function
  End If
  
  'Look at each file that was unzipped, and check to see if any of them are
  'VB files - if not, return nothing
  For Each lFile In FSO.GetFolder(lPath).Files
    If IsVBFile(BuildFileType(FSO.GetFile(lFile).Name).Ext) Then
      UnzipFile = lPath
      RZip.FileName = ""
      Exit Function
    End If
  Next lFile
  
  RZip.FileName = ""
      
End Function

'Removes a node from the tree
Private Function RemoveNode(TV As TreeView, tKey As String)

  If IsRootNode(tKey) Then
    RemoveRootNode tKey
    ParentFiles.Remove tKey
  End If
  
  TV.Nodes.Remove tKey
  
End Function

'Adds a new node to the tree
Private Function AddNewNode(TV As TreeView, tPKey As String, tKey As String, tName As String, tIconKey As String, tExpanded As Boolean)

  'If the parent key does not exist then it is a root node
  If tPKey = "" Then
    AddRootNode tKey
    ParentFiles.Add tKey
  Else
    ParentFiles.Item(tPKey).Add tKey
  End If
  
  'Add the node
  If tPKey = "" Then
    TV.Nodes.Add(, tvwChild, tKey, tName, tIconKey).Sorted = True
  Else
    TV.Nodes.Add(tPKey, tvwChild, tKey, tName, tIconKey).Sorted = True
  End If
  TV.Nodes(tKey).Expanded = True
  
  If tKey = "Shared" Then TV.Nodes(tKey).SelectedImage = "OpenFolder"

End Function

'Adds a file to the project tree and other required areas
Private Function AddFileToTree(TV As TreeView, lFullFile As String, tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary, LV As ListView, RZip As RichsoftVBZip, Optional ParentNode As String = "")

  Dim lFile As typFile
  Dim lZipPath As String
  Dim vFile As Variant
  Dim FileHasBeenManaged As Boolean
  Dim lImage As String
  
  lFile = BuildFileType(lFullFile)
  
  'Handle Zip Stuff
  If IsZipFile(UCase$(lFile.Ext)) Then
    'If the ZIP file has already been added then get out
    If KeyExists(TV, lFullFile) Then Exit Function
    'Check to see if the zip file was unzipped and if it contains any VB files
    lZipPath = UnzipFile(TV, lFullFile, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, RZip)
    'If no zip file path was returned then the file did not contain any VB files
    If lZipPath = "" Then Exit Function
    'Otherwise, add the zip file to the tree
    AddNewNode TV, "", lFullFile, lFile.Name & "." & lFile.Ext, UCase$(lFile.Ext), True
    'Attempt to add each file from the zip file to the project tree
    For Each vFile In FSO.GetFolder(lZipPath).Files
      AddFileToTree TV, CStr(vFile), tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, RZip, lFullFile
    Next vFile
    'This is all that we do with Zip files
    Exit Function
  End If
  
  FileHasBeenManaged = False
  
  If Not IsVBFile(UCase$(lFile.Ext)) Then Exit Function
  
  'Check to see if the key already exist
  If Not KeyExists(TV, lFullFile) Then
    AddNewNode TV, ParentNode, lFullFile, lFile.vbName & " (" & lFile.Name & "." & lFile.Ext & ")", UCase$(lFile.Ext), True
  Else
    'Check to see if this is a root node
    If ParentNode <> "" Then
      'Check to see if the existing node is a root node
      If IsRootNode(lFullFile) Then
        FileHasBeenManaged = True
        lImage = TV.Nodes(lFullFile).Image
        RemoveNode TV, lFullFile
        AddNewNode TV, ParentNode, lFullFile, lFile.vbName & " (" & lFile.Name & "." & lFile.Ext & ")", lImage, True
      Else
        'Check to see if it is the same file that we are referencing
        If ParentNode <> TV.Nodes(lFullFile).Parent.Key Then
          'Check to see if the existing node is contained within a zip file
          If AnyParentIsZip(TV, lFullFile) Then
            'If any elder is a Zip, then organize this project as needed
            If Not IsRootNode(ParentNode) Then
              If AnyParentIsZip(TV, ParentNode) Then
                FileHasBeenManaged = True
                lImage = TV.Nodes(lFullFile).Image
                RemoveNode TV, lFullFile
                AddNewNode TV, ParentNode, lFullFile, lFile.vbName & " (" & lFile.Name & "." & lFile.Ext & ")", lImage, True
              End If
            End If
          Else
            'Check to see is my parent is a zip file
            If UCase$(Right$(ParentNode, 3)) = "ZIP" Then
              FileHasBeenManaged = True
              lImage = TV.Nodes(lFullFile).Image
              RemoveNode TV, lFullFile
              AddNewNode TV, ParentNode, lFullFile, lFile.vbName & " (" & lFile.Name & "." & lFile.Ext & ")", lImage, True
            Else
              'Checks to see if the "Shared Folder" exists
              If Not IsRootNode("Shared") Then AddNewNode TV, "", "Shared", "Shared Files", "ClosedFolder", True
              'Adds the key
              If IsNotShared(TV, lFullFile) Then
                FileHasBeenManaged = True
                lImage = TV.Nodes(lFullFile).Image
                RemoveNode TV, lFullFile
                AddNewNode TV, "Shared", lFullFile, lFile.vbName & " (" & lFile.Name & "." & lFile.Ext & ")", lImage, True
              End If
            End If
          End If
        End If
      End If
    End If
  End If
  
  'If this file is a project or a project group then enumerate its' children
  If UCase$(lFile.Ext) = "VBP" Or UCase$(lFile.Ext) = "VBG" Then
    EnumerateProjectChildren TV, lFile, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, RZip
  End If
  
  If Not FileHasBeenManaged And FlySearch Then
    Call SearchFile(tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, TV.Nodes(lFullFile))
  End If

End Function

'Check to see if the file is shared by more that one project
Private Function IsNotShared(TV As TreeView, tKey As String) As Boolean

  If TV.Nodes(tKey).Root.Key = "Shared" Then IsNotShared = False Else IsNotShared = True
  
End Function

'Checks to see if any level parent of the node is a zip file
Private Function AnyParentIsZip(TV As TreeView, tKey As String) As Boolean

  Dim tPKey As String
  Dim tRKey As String
  
  'Set default
  AnyParentIsZip = False
  
  'Set the root key
  tRKey = TV.Nodes(tKey).Root.Key
  'Set the parent key (therefore if the passed key is root, the loop will be ignored)
  tPKey = tKey
  
  Do Until tPKey = tRKey
    'Get the parent key
    tPKey = TV.Nodes(tPKey).Parent.Key
    'Check to see if it is a zip file
    If UCase$(Right$(tPKey, 3)) = "ZIP" Then
      AnyParentIsZip = True
      Exit Function
    End If
  Loop

End Function

'Gets the Forms, Classes, etc. in the project
Private Function EnumerateProjectChildren(TV As TreeView, lFile As typFile, tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary, LV As ListView, RZip As RichsoftVBZip)

  Dim tKey As String
  Dim tChildKey As String
  Dim FreeNum As Integer
  Dim dummy As String
  Dim d As String
  
  On Error GoTo Err_Handle
  
  'Define the key
  tKey = lFile.FilePath & lFile.Name & "." & lFile.Ext
  'Add this key as a parent
  ParentFiles.Add tKey
  
  'Get a free filenumber
  FreeNum = FreeFile
  
  'Open the file
  Open tKey For Input As #FreeNum
    'Get one line (always garbage in VBP & VBG files)
    Line Input #FreeNum, dummy
    'Loop until we hit the end of the file
    Do Until EOF(FreeNum)
      
      Line Input #FreeNum, dummy
      
      d = ""
      
      'Search for project/project group attributes
      If UCase(lFile.Ext) = "VBG" Then
        d = AttemptExtractProjects(dummy)
      ElseIf UCase(lFile.Ext) = "VBP" Then
        d = AttemptExtractProjectFiles(dummy)
      End If
      
      'If a child was found then
      If d <> "" Then
        'Generate the key of the new child
        tChildKey = GenerateChildPath(lFile.FilePath, d)
        'Add the new child
        If FSO.FileExists(tChildKey) Then
          ParentFiles.Item(tKey).Add tChildKey
          AddFileToTree TV, tChildKey, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, RZip, tKey
        End If
      End If
      
    Loop
    
Err_Handle:
  Close #FreeNum
  Exit Function
  'We need to put resume next here for some versions of VB6
  Resume Next

End Function

'This generates the absolute path of a child based on the path of a the
'parent using relative, or absolute paths.
Private Function GenerateChildPath(pPath As String, cPath As String) As String

  Dim x As Integer
  Dim y As Integer
  Dim d1 As String
  Dim d2 As String
  
  'If there are no \'s then the child is in the same directory
  'as the parent
  If InStr(1, cPath, "\") = 0 Then
    GenerateChildPath = pPath & cPath
  'Otherwise...
  Else
    'If the first three characters of the string are ..\ then
    'this is relative addressing
    If Left$(cPath, 3) = "..\" Then
      y = 0
      d1 = cPath
      'strip off the ..\'s and count how many this was
      Do Until Left$(d1, 3) <> "..\"
        d1 = Right$(d1, Len(d1) - 3)
        y = y + 1
      Loop
      
      d2 = pPath
      
      'Strip the same number of child directories off of the end
      'of the parent directory
      For x = 1 To y
        d2 = StripLastDirectory(d2)
      Next x
      
      'Fuse if all back together
      GenerateChildPath = d2 & d1
      
    'If there is a : then we have been passed an absolute path
    ElseIf Mid$(cPath, 2, 1) = ":" Then
      GenerateChildPath = cPath
    Else
      GenerateChildPath = pPath & cPath
    End If
  End If

End Function

'Removes the last directory in a directory listing
Private Function StripLastDirectory(lPath As String) As String

  Dim y1 As Integer
  Dim y2 As Integer
  
  'Strip the last \ if needed
  If Right$(lPath, 1) = "\" Then lPath = Left$(lPath, Len(lPath) - 1)
  
  y1 = -1
  
  'Count the number of directory structures
  Do Until y1 = 0
    If y1 = -1 Then y1 = 0
    y2 = y1
  
    y1 = InStr(y1 + 1, lPath, "\")
    
  Loop
  
  'Remove the last one unless we are down to c: then add the \ back on
  If y2 = 0 Then
    StripLastDirectory = lPath & "\"
  Else
    StripLastDirectory = Left$(lPath, y2)
  End If

End Function

'Removes a root node from the array
Public Function RemoveRootNode(tKey As String)

  Dim x As Integer
 
  'Look at each element in the array
  'and remove the selected key
  For x = LBound(RootNodes) To UBound(RootNodes)
    If RootNodes(x) = tKey Then
      RootNodes(x) = ""
      Exit Function
    End If
  Next x
  
End Function

'Adds an element to the root node array
Private Function AddRootNode(tKey As String)

  Dim x As Integer
  Dim FoundSpace As Boolean
      
  'Since we never shrink the array, we need to look for spaces
  'where we have previously removed data
  FoundSpace = False
  
  'Look at each element in the array
  For x = LBound(RootNodes) To UBound(RootNodes)
    If RootNodes(x) = "" Then
      FoundSpace = True
      Exit For
    End If
  Next x
  
  'If we didn't find an space in the array then we need to make
  ' it bigger
  If Not FoundSpace Then
    x = UBound(RootNodes) + 1
    ReDim Preserve RootNodes(x)
  End If
  
  'Set the array element
  RootNodes(x) = tKey

End Function

'Determines if the key is in the root of the tree
Public Function IsRootNode(tKey As String) As Boolean

  Dim x As Integer
  
  'Initialize the variable
  IsRootNode = False
  
  'Look through each element and see if the node exists
  For x = LBound(RootNodes) To UBound(RootNodes)
  
    If RootNodes(x) = tKey Then
      IsRootNode = True
      Exit Function
    End If
    
  Next x

End Function

'Determines if the extension is a Zip extension
Private Function IsZipFile(strExt As String) As Boolean
  
  If strExt = "ZIP" Then IsZipFile = True Else IsZipFile = False

End Function

'Determines if the extension is a VB extension
Private Function IsVBFile(strExt As String) As Boolean

  Dim x As Integer
  
  'Initialize the variable
  IsVBFile = False
  
  'Look through the list of valid extensions and see if ours exist
  For x = LBound(ValidVBFiles) To UBound(ValidVBFiles)
    
    If ValidVBFiles(x) = UCase$(strExt) Then
      IsVBFile = True
      Exit Function
    End If
  
  Next x

End Function

'Checks to see if a key is being used in a Treeview control
Private Function KeyExists(TV As TreeView, tKey As String) As Boolean

  Dim x As Integer
  
  'Initialize the variable
  KeyExists = False
  
  'If there are no nodes then no keys exist
  If TV.Nodes.Count = 0 Then Exit Function
  
  'Look through each element and see if the key is being used
  For x = 1 To TV.Nodes.Count
    If TV.Nodes(x).Key = tKey Then
      KeyExists = True
      Exit Function
    End If
  Next x

End Function

'Builds a file type based on a filename
Public Function BuildFileType(lFullFile As String) As typFile

  Dim y1 As Integer
  Dim y2 As Integer
  Dim lFileName As String
  
  'Gets the name of the file
  lFileName = FSO.GetFile(lFullFile).Name
  
  'Looks for the last period in the file name
  y1 = -1
  Do Until y1 = 0
    y2 = y1
    If y1 = -1 Then y1 = 0
    y1 = InStr(y1 + 1, lFileName, ".")
  Loop
  
  'Set the file type and the file name
  If y2 = -1 Then
    BuildFileType.Ext = ""
    BuildFileType.Name = lFileName
  Else
    BuildFileType.Ext = Right$(lFileName, Len(lFileName) - y2)
    BuildFileType.Name = Left$(lFileName, y2 - 1)
  End If
   
  'Gets the Windows file type definition
  BuildFileType.Type = FSO.GetFile(lFullFile).Type
  'Gets the path of the file
  BuildFileType.FilePath = Replace(Left$(lFullFile, Len(lFullFile) - Len(lFileName)) & "\", "\\", "\")
  'Gets the VB friendly file name
  BuildFileType.vbName = GetVBName(lFullFile)
  'Gets the file size
  BuildFileType.Size = Int(FSO.GetFile(lFullFile).Size / 10.24) / 100
  'Gets the full file name
  BuildFileType.FullName = lFullFile
      
End Function

'Gets the VB friendly filename from the file
Private Function GetVBName(lFullFile As String) As String

  On Error GoTo Err_Handle
  
  Dim FreeNum As Integer
  Dim dummy As String
  Dim d As String
  
  'Get an available file number
  FreeNum = FreeFile
  
  'Open the file for reading
  Open lFullFile For Input As #FreeNum
    'Look through each line of the file
    Do Until EOF(FreeNum)
      Line Input #FreeNum, dummy
      'Try to get the file name
      d = AttemptExtractName(dummy, "VersionFileDescription=")
      If d <> "" Then GoTo Err_Handle
      
      d = AttemptExtractName(dummy, "Attribute VB_Name = ")
      If d <> "" Then GoTo Err_Handle
      
      d = AttemptExtractName(dummy, "Name=")
      If d <> "" Then GoTo Err_Handle
      
      d = AttemptExtractName(dummy, "Begin VB.Form")
      If d <> "" Then GoTo Err_Handle
       
    Loop
  
Err_Handle:
  Close FreeNum
  GetVBName = d
  Exit Function
  Resume Next

End Function

'Tries to get the projects from a project group
Private Function AttemptExtractProjects(SearchString As String) As String

  Dim d As String
  'Looks at each option and attempts to extract the project file name
  d = AttemptExtractName(SearchString, "StartupProject=")
  If d = "" Then d = AttemptExtractName(SearchString, "Project=")
    
  AttemptExtractProjects = d

End Function

'Tries to get the files (forms, modules, etc.) from a project
Private Function AttemptExtractProjectFiles(SearchString As String) As String

  Dim d As String
  Dim y As Integer
  
  'Looks at each option and tries to get the filename
  d = AttemptExtractName(SearchString, "Form=")
  If d = "" Then
    d = AttemptExtractName(SearchString, "Module=")
    If d = "" Then
      d = AttemptExtractName(SearchString, "UserControl=")
      If d = "" Then
        d = AttemptExtractName(SearchString, "Class=")
        If d = "" Then
          d = AttemptExtractName(SearchString, "PropertyPage=")
        End If
      End If
    End If
  End If
    
  'Gets rid of any un-needed information
  If Len(d) > 0 Then
    y = InStr(1, d, ";")
    If y <> 0 Then d = Trim(Right(d, Len(d) - y))
  End If
  
  AttemptExtractProjectFiles = d

End Function

'Attempts to extract a VB friendly name from a line
Private Function AttemptExtractName(SearchString As String, SearchCriteria As String) As String

  Dim d As String
  
  'Looks at the first few characters to see if it matches the search criteria
  If Len(SearchString) > Len(SearchCriteria) Then
    If Left(SearchString, Len(SearchCriteria)) = SearchCriteria Then _
        d = Mid$(SearchString, Len(SearchCriteria) + 1, Len(SearchString) - (Len(SearchCriteria))) _
        Else: d = ""
  End If
  
  'Gets rid of extra spaces
  d = Trim(d)
  
  'Trims off the " from each end if they are there
  If Len(d) > 0 Then
    If Left(d, 1) = vbQ Then d = Right(d, Len(d) - 1)
  End If
  
  If Len(d) > 0 Then
    If Right(d, 1) = vbQ Then d = Left(d, Len(d) - 1)
  End If
   
  AttemptExtractName = d

End Function

'Sets the frame that should be displayed when a tab is clicked (frmMain)
Public Function SetTabFrame(Index As Integer, FrameState As Boolean)

  Select Case Index
    Case 1
      frmMain.fraProjectList.Visible = FrameState
    Case 2
      frmMain.fraSearchList.Visible = FrameState
  End Select

End Function

'Initializes various items
Public Function InitializeProject()
  
  SeverityColor(0) = 2
  SeverityColor(1) = 6
  SeverityColor(2) = 4
  
  SeverityDescription(0) = "Low"
  SeverityDescription(1) = "Normal"
  SeverityDescription(2) = "High"
  
  ValidVBFiles(0) = "VBP"
  ValidVBFiles(1) = "FRM"
  ValidVBFiles(2) = "BAS"
  ValidVBFiles(3) = "CLS"
  ValidVBFiles(4) = "CTL"
  ValidVBFiles(5) = "PAG"
  ValidVBFiles(6) = "VBG"
  
  ProcedureNames(0) = " Sub "
  ProcedureNames(1) = " Function "
  ProcedureNames(2) = " Property "
  
  ReDim RootNodes(1)
    
  frmMain.cmbSeverity.ListIndex = 1
  
  frmMain.picProgress.Width = 0
    
  frmMain.PT.Nodes.Add(, , , "Drag and Drop your").ForeColor = QBColor(8)
  frmMain.PT.Nodes.Add(, , , "project files here.").ForeColor = QBColor(8)
  
  blnProjectTreeInitialized = False
    
  vbQ = Chr(34)
  SearchPercent = 0
  
  On Error Resume Next
  
  FlySearch = GetSetting(AppTitle, "Options", "FlySearch", True)
  SaveDictionary = Int(GetSetting(AppTitle, "Options", "SaveDictionary", 1))
  IgnoreComments = GetSetting(AppTitle, "Options", "IgnoreComments", True)
  RemoveZips = GetSetting(AppTitle, "Options", "RemoveZips", True)

End Function

'Generates a new key for a given listview control
Public Function GenerateKey(LV As ListView) As String

  Dim x As Integer
  Dim blnKeyIsUnique As Boolean
  Dim tKey As String
  
  Randomize
  
  blnKeyIsUnique = False
  
  Do Until blnKeyIsUnique
  
    tKey = ""
    
    'Generates a 10 character key using uppercase letters
    For x = 1 To 10
      tKey = tKey & Chr(Int(Rnd * 26) + 65)
    Next x
    
    blnKeyIsUnique = True
    
    'Checks each item in the listview to make sure that our key is unique
    If LV.ListItems.Count <> 0 Then
      For x = 1 To LV.ListItems.Count
        If LV.ListItems(x).Key = tKey Then blnKeyIsUnique = False
      Next x
    End If
  
  Loop

  GenerateKey = tKey

End Function

'Exports our search dictionary to a file
Public Function ExportDictionary(DestFile As String, SD As clsDictionary)

  On Error GoTo Err_Handle
  
  Dim FreeNum As Integer
  Dim x As Integer
  Dim tOut As String
  
  FreeNum = FreeFile
  
  Open DestFile For Output As #FreeNum
    If SD.Count <> 0 Then
      For x = 1 To SD.Count
        tOut = "PHRASE|" & Replace(SD.Item(x).Phrase, "|", "||") & "|"
        tOut = tOut & "SEVERITY|" & SD.Item(x).Severity & "|"
        tOut = tOut & "DESCRIPTION|" & Replace(SD.Item(x).Description, "|", "||") & "|"
        
        Print #FreeNum, tOut
      
      Next x
    End If
      
Err_Handle:
  Close FreeNum
  Exit Function
  Resume Next

End Function

'Imports our search dictionary from a file
Public Function ImportDictionary(SourceFile As String, SD As clsDictionary, LV As ListView)

  On Error GoTo Err_Handle
  
  Dim FreeNum As Integer
  Dim dummy As String
  Dim d As String
  Dim tPhrase As String
  Dim tDescription As String
  Dim tSeverity As Integer
  Dim tLineValid As Boolean
  
  If Not FSO.FileExists(SourceFile) Then Exit Function
  
  FreeNum = FreeFile
  
  Open SourceFile For Input As #FreeNum
  
    Do Until EOF(FreeNum)
      Line Input #FreeNum, dummy
      
      Call ParseImportString(dummy, tPhrase, tDescription, tSeverity, tLineValid)
      If tLineValid Then SD.Add GenerateKey(LV), tPhrase, tSeverity, tDescription
        
    Loop
    
Err_Handle:
  Close FreeNum
  
  DictionaryChanged = True
  
  Exit Function
  Resume Next
  
End Function

'Parses the import string and sets each variable
'(Ummmmm ..... I don't feel like commenting this ...... better luck next version)
Private Sub ParseImportString(tParse As String, tPhrase As String, tDescription As String, tSeverity As Integer, tLineValid As Boolean)

  Dim y1 As Integer
  Dim y2 As Integer
  Dim d As String
  Dim SingleFound As Boolean
  
  tLineValid = False
  y1 = -1
  
  If Len(tParse) < 4 Then Exit Sub
  If Right$(tParse, 1) <> "|" Then Exit Sub
  
  Do Until y1 = Len(tParse)
    If y1 = -1 Then y1 = 0
    y2 = y1 + 1
    
    y1 = InStr(y1 + 1, tParse, "|")
    
    SingleFound = False
    
    If y1 + 1 = InStr(y1 + 1, tParse, "|") Then
      
      Do Until SingleFound
      
        SingleFound = True
      
        Do Until y1 + 1 <> InStr(y1 + 1, tParse, "|") Or y1 = Len(tParse)
          y1 = y1 + 1
        Loop
      
        If y1 <> Len(tParse) Then
          y1 = y1 + 1
          y1 = InStr(y1 + 1, tParse, "|")
          If y1 + 1 = InStr(y1 + 1, tParse, "|") Then SingleFound = False
        Else
          SingleFound = True
        End If
                
      Loop
            
    End If
          
    If d = "" Then
      d = Mid$(tParse, y2, y1 - y2)
    Else
      Select Case d
        Case "PHRASE"
          tPhrase = Replace(Mid$(tParse, y2, y1 - y2), "||", "|")
          If tPhrase = "" Then Exit Sub
        Case "DESCRIPTION"
          tDescription = Replace(Mid$(tParse, y2, y1 - y2), "||", "|")
        Case "SEVERITY"
          d = Replace(Mid$(tParse, y2, y1 - y2), "||", "|")
          If IsNumeric(d) Then tSeverity = Int(d) Else Exit Sub
      End Select
      d = ""
    End If
        
  Loop
    
  tLineValid = True

End Sub

'Searches a file for violations based on the search dictionary
Public Function SearchFile(tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary, LV As ListView, Node As MSComctlLib.Node)

  Dim x As Integer
  Dim lFile As String
  Dim FreeNum As Integer
  Dim dummy As String
  Dim lFileSize As Long
  
  'The filename is the key
  lFile = Node.Key
  
  'If the file doesn't exist, get out
  If Not FSO.FileExists(lFile) Then Exit Function
  'Gets the filesize
  lFileSize = FSO.GetFile(lFile).Size
  
  On Error GoTo Err_Handle
  
  FreeNum = FreeFile
  
  'Sets environment variables
  tProgressCaption.Caption = "Searching " & FSO.GetFile(lFile).Name
  'Refreshes the progress bar
  tProgressBar.Refresh
  'Initializes variables
  CurrentProcedure = "<Unknown>"
  FoundProcedure = False
  ViolationFound = False
  SearchPercent = 0
  
  'Opens the file
  Open lFile For Input As #FreeNum
    
    Do Until EOF(FreeNum)
    
      'Gets a line
      Line Input #FreeNum, dummy
      'Sets the current search percent (1 character = 1 byte)
      SearchPercent = SearchPercent + (Len(dummy) / lFileSize)
      'Sets the progress indicator
      tProgressBar.Width = Int(SearchPercent * tProgressBarContainer.Width) + tProgressBarOffset
      tProgressBar.Refresh
      'Some function calls
      Call SetCurrentProcedure(dummy)
      
      Call SearchLine(dummy, SD, LV, Node, lFile)
      
    Loop
    
Err_Handle:
  
  Close #FreeNum
  'Resets the form
  tProgressCaption.Caption = "Not Currently Searching"
  SearchPercent = 0
  tProgressBar.Width = 0
  CurrentProcedure = "<Unknown>"
  FoundProcedure = False
  
  If Len(Node.Image) <= 4 And Node.Image <> "ZIP" Then
    If ViolationFound Then Node.Image = "CROSS" & Node.Image _
        Else Node.Image = "CHECK" & Node.Image
  End If
  
  Exit Function
  Resume Next
  
End Function

'Resets the current procedure that we are scanning
Private Function SetCurrentProcedure(tLine As String) As Boolean

  Dim x As Integer
  Dim d As String
    
  SetCurrentProcedure = False
  
  If Len(tLine) < 2 Then Exit Function
  
  SetCurrentProcedure = True
  'If we havn't found a procedure before then we are dealing with a property
  If Not FoundProcedure Then
    If InStr(1, Trim(tLine), "Begin ") = 1 Then
      d = Right$(Trim(tLine), Len(Trim(tLine)) - 6)
      d = Right$(d, Len(d) - InStr(1, d, "."))
      x = InStr(1, d, " ")
      CurrentProcedure = "Property within " & Right$(d, Len(d) - x) & " - " & Left$(d, x - 1) & " Control"
      Exit Function
    End If
  End If
  
  'Look through the line for valid procedure definition names and
  'make sure that we are calling the procedure and not ending it
  For x = LBound(ProcedureNames) To UBound(ProcedureNames)
  
    If InStr(1, tLine, ProcedureNames(x)) <> 0 Then
      If InStr(1, tLine, "End " & Trim(ProcedureNames(x))) = 0 Then
        If InStr(1, tLine, "Exit " & Trim(ProcedureNames(x))) = 0 Then
          FoundProcedure = True
          CurrentProcedure = tLine
          Exit Function
        End If
      End If
    End If
    
  Next x
  
  SetCurrentProcedure = False
      
End Function

'Searches the given line for a violation
Private Function SearchLine(tLine As String, SD As clsDictionary, LV As ListView, Node As MSComctlLib.Node, lFile As String) As Boolean

  Dim x As Integer
  Dim y1 As Integer
  Dim y2 As Integer
  Dim tKey As String
  Dim AddLinetoList As Boolean
    
  SearchLine = False
  
  If Len(tLine) = 0 Then Exit Function
  
  For x = 1 To SD.Count
  
    y1 = InStr(1, tLine, SD.Item(x).Phrase, vbTextCompare)
    y2 = InStr(1, tLine, "'")
    AddLinetoList = False
  
    If y1 <> 0 Then
      If IgnoreComments Then
        If y2 > y1 Or y2 = 0 Then AddLinetoList = True
      Else
        AddLinetoList = True
      End If
    End If
      
    If AddLinetoList Then
    
      'If we've found one then add the data to the list view
      tKey = GenerateKey(LV)
      LV.ListItems.Add , tKey, FSO.GetFile(lFile).Name & " - " & lFile
      LV.ListItems(tKey).ListSubItems.Add 1, , SD.Item(x).Phrase
      LV.ListItems(tKey).ListSubItems.Add(2, , SeverityDescription(SD.Item(x).Severity)).ForeColor = QBColor(SeverityColor(SD.Item(x).Severity))
      LV.ListItems(tKey).ListSubItems.Add 3, , Trim(tLine)
      LV.ListItems(tKey).ListSubItems.Add 4, , CurrentProcedure
      LV.ListItems(tKey).ListSubItems.Add 5, , SD.Item(x).Description
      SearchLine = True
      ViolationFound = True
                  
      LV.Refresh
                  
    End If
        
  Next x

End Function

'Saves the search dictionary to the registry
Public Function SaveSearchDictionary(SD As clsDictionary)

  Dim x As Integer
  
  SaveSetting AppTitle, "Dictionary", "Total Definitions", SD.Count
  
  If SD.Count = 0 Then Exit Function
  
  For x = 1 To SD.Count
  
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Phrase", SD.Item(x).Phrase
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Severity", SD.Item(x).Severity
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Description", SD.Item(x).Description
           
  Next x

  DictionaryChanged = False

End Function

'Loads the search dictionary from the registry
Public Function LoadSearchDictionary(SD As clsDictionary, LV As ListView)

  Dim x As Integer
  Dim d As String
  Dim tCount As Integer
  Dim tPhrase As String
  Dim tSeverity As Integer
  Dim tDescription As String
  
  d = GetSetting(AppTitle, "Dictionary", "Total Definitions", 0)
  If IsNumeric(d) Then tCount = Int(d) Else Exit Function
    
  If tCount <= 0 Then Exit Function
  
  For x = 1 To tCount
  
    tPhrase = GetSetting(AppTitle, "Dictionary", "Def" & x & "_Phrase", "")
    If tPhrase = "" Then Exit Function
    
    d = GetSetting(AppTitle, "Dictionary", "Def" & x & "_Severity", 1)
    If IsNumeric(d) Then tSeverity = Int(d) Else Exit Function
    If tSeverity > 2 Then tSeverity = 2
    If tSeverity < 0 Then tSeverity = 0
    
    tDescription = GetSetting(AppTitle, "Dictionary", "Def" & x & "_Description", "")
    
    SD.Add GenerateKey(LV), tPhrase, tSeverity, tDescription
    
  Next x
  
  DictionaryChanged = False

End Function

'Triggered when we are dragging an OLE item over the form
Public Function PublicOLEDragOver(tForm As Form)

  'If the project search tab is not displayed then display it
  If tForm.TabStrip1.SelectedItem.Index <> 1 Then tForm.TabStrip1.Tabs(1).Selected = True
  'If the project is minimized then fix that
  If tForm.WindowState = vbMinimized Then tForm.WindowState = vbNormal
  'If this is not the active window the active it
  If GetForegroundWindow <> tForm.hwnd Then tForm.SetFocus
  

End Function

'Triggered when a file(s) is/are dropped onto our form
Public Function PublicOLEDragDrop(Data As Object, TV As TreeView, tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary, LV As ListView, IL As ImageList, RZip As RichsoftVBZip)
  
  Dim z As Integer
  Dim lFile As String

  'Initialize the tree if it isn't
  If Not blnProjectTreeInitialized Then
    blnProjectTreeInitialized = True
    TV.ImageList = IL
    TV.Nodes.Clear
  End If

  'Look at each file
  For z = 1 To Data.Files.Count

    lFile = Data.Files.Item(z)
    If FSO.FileExists(lFile) Then AddFileToTree TV, lFile, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, RZip

  Next z

End Function

'Triggered when we click a column header to organize it
Public Function SortListView(LV As ListView, ColumnHeader As MSComctlLib.ColumnHeader)

  If LV.SortKey <> ColumnHeader.Index - 1 Then
    LV.SortOrder = lvwAscending
    LV.SortKey = ColumnHeader.Index - 1
  Else
    If LV.SortOrder = lvwAscending Then
      LV.SortOrder = lvwDescending
    Else
      LV.SortOrder = lvwAscending
    End If
  End If
    
  LV.Sorted = True
  LV.Refresh
  LV.Sorted = False

End Function

'Triggered at startup
Public Function SetupWizard(tForm As Form, SD As clsDictionary, LV As ListView)

  Dim d As String
  Dim NoPreviousVersion As Boolean
  Dim q As Variant
  Const HighestQStep = "4"

  NoPreviousVersion = False

  'Setup Wizard
  'Checks to see if this is the first time that the app is run
  d = GetSetting(AppTitle, "FirstRun", "HasRun", "-1")
  
  If d = "-1" Then d = GetSetting("Source Search 1.0", "FirstRun", "HasRun", "0")
    
  If d <> "0" And d <> HighestQStep Then
    MsgBox "Because you are running an updated version of Source Search, we will need to ask you a couple of questions related to new features.", , "First Run Wizard"
  End If
  
  If d = "0" Then
    'Update the setting
    SaveSetting AppTitle, "FirstRun", "HasRun", HighestQStep
    d = HighestQStep
    
    NoPreviousVersion = True
    
    'Asks first run questions regarding options and importing a dictionary
    MsgBox "Because this is your first time running this application, we will need to setup a couple of options.", vbOKOnly, "First Run Wizard"
    
    frmOptions.Show vbModal, tForm
    
    q = MsgBox("A Search Dictionary is a list of words to search for, along with some definitions. You can choose to import the list that shipped with Source Search, or you can build your own. You can always import a new dictionary later. Would you like to import Source Search's startup dictionary?", vbYesNo + vbQuestion, "First Run Wizard")
    
    If q = vbYes Then
      'Checks to see the default dictionary exists and imports it.
      'If it doesn't exist then the user is propted to select one.
      If FSO.FileExists(Replace(App.Path & "\", "\\", "\") & "Startup Dictionary.sdd") Then
        ImportDictionary Replace(App.Path & "\", "\\", "\") & "Startup Dictionary.sdd", SD, LV
      Else
        
        'Clears the dialog filename
        tForm.CommDiag.FileName = ""
        tForm.CommDiag.ShowOpen
        'If the dialog filename = "" then the user clicked cancel
        If tForm.CommDiag.FileName <> "" Then
          'Calls the global import function
          ImportDictionary tForm.CommDiag.FileName, SD, LV
        End If
        
      End If
    End If
  End If
  
  If Not NoPreviousVersion Then
  
    If d = "1" Then
      
      q = MsgBox("An author's code comments can give insight to the intent of the code, however, these comments can include harmless words that would be considered malicious if the were executed in code. Would you like to ignore any commented code?", vbYesNo + vbQuestion, "First Run Wizard")
      If q = vbYes Then IgnoreComments = True Else IgnoreComments = False
      
      'Update the setting
      SaveSetting AppTitle, "FirstRun", "HasRun", "2"
      d = "2"
    
    End If
      
    If d = "2" Then
      
      MsgBox "To use the Zip features of Source Search, you will need to copy:" & vbCrLf & vbCrLf & vbTab & "Unzdll.dll" & vbCrLf & vbTab & "Zipdll.dll" & vbCrLf & vbTab & "Zipit.dll" & vbCrLf & vbCrLf & "to your ...\Windows\System32\ directory.", , "First Run Wizard"
      
      q = MsgBox("Source Search can search zip files that contain VB files for malicious code. Would you like Source Search to automatically clean up its' temporary files when it closes? (This feature is highly recommended as a large amount of hard drive space can be used up quickly.)", vbYesNo + vbQuestion, "First Run Wizard")
      
      If q = vbYes Then RemoveZips = True Else RemoveZips = False
          
      If FSO.FileExists(Replace(App.Path & "\", "\\", "\") & "Dictionary Update 2.0.1.sdd") Then
        q = MsgBox("Would you like to add some new dictionary definitions to your existing dictionary?", vbYesNo + vbQuestion, "First Run Wizard")
        If q = vbYes Then ImportDictionary Replace(App.Path & "\", "\\", "\") & "Dictionary Update 2.0.1.sdd", SD, LV
      End If
      
      'Update the setting
      SaveSetting AppTitle, "FirstRun", "HasRun", "3"
      d = "3"
    
    End If
    
    If d = "3" Then
    
      q = MsgBox("Source Search dictionary and setting storeage has also been modified. Would you like Source Search to upgrade this feature? (You will not have access to your search dictionary or your personal options if you choose no.)", vbYesNo + vbQuestion, "First Run Wizard")
    
      If q = vbYes Then
        SetupWizard_UpgradeOptions
      Else
        q = MsgBox("Would you like Source Search to ask you this again next time you start the application? (If you choose no, your search dictionary and personal options will be perminently lost.)", vbYesNo + vbQuestion, "First Run Wizard")
      End If
                
      If q = vbNo Then
        SaveSetting AppTitle, "FirstRun", "HasRun", "4"
        d = "4"
      End If
      
    End If
  
  End If
  
End Function

'Triggered when the user chooses to view a node's information
Public Function ShowNodeInfo(tForm As Form, LV As ListView, TV As TreeView)

  'If the tree has not been initialized or is empty, get out
  If Not blnProjectTreeInitialized Or TV.Nodes.Count < 1 Then Exit Function
    
  'Load the form
  Load frmNodeInfo
  
  'Set some basic values
  frmNodeInfo.txtDisplayName.Text = TV.SelectedItem.Text
  frmNodeInfo.txtPath.Text = TV.SelectedItem.Key
        
  'Set Advanced options through routines
  ShowNodeInfo_Relations TV
  ShowNodeInfo_TreeView TV
  ShowNodeInfo_Violations TV, LV
  
  'Show the form modally
  frmNodeInfo.Show vbModal, tForm

End Function

'Show the relationship information
Private Function ShowNodeInfo_Relations(TV As TreeView)

  Dim x As Integer
  Dim y As Integer
  Dim ChildListClear As Boolean
  Dim ParentListClear As Boolean

  'Sets default values
  ChildListClear = False
  ParentListClear = False
  
  'If there are any parent files
  If ParentFiles.Count > 0 Then
    'Look at each one
    For x = 1 To ParentFiles.Count
      'If there are any Children in this parent
      If ParentFiles.Item(x).Count > 0 Then
        'Look at each one
        For y = 1 To ParentFiles.Item(x).Count
          'If the selected item (in the project tree) is a child,
          If ParentFiles.Item(x).Item(y).Path = TV.SelectedItem.Key Then
            If Not ParentListClear Then
              frmNodeInfo.lstParents.Clear
              ParentListClear = True
            End If
            'Add it to the parent list
            frmNodeInfo.lstParents.AddItem TV.Nodes(ParentFiles.Item(x).Path).Text
          End If
          
          'If the selected item is a parent
          If ParentFiles.Item(x).Path = TV.SelectedItem.Key Then
            If Not ChildListClear Then
              frmNodeInfo.lstChildren.Clear
              ChildListClear = True
            End If
            'Add this child
            frmNodeInfo.lstChildren.AddItem TV.Nodes(ParentFiles.Item(x).Item(y).Path).Text
          End If
          
         Next y
      End If
    Next x
  End If


End Function

'Show the tree of immediate children in the info tree
Private Function ShowNodeInfo_TreeView(TV As TreeView)

  Dim x As Integer
  
  'Add the top node
  frmNodeInfo.TV.Nodes.Add(, , TV.SelectedItem.Key, TV.SelectedItem.Text, TV.SelectedItem.Image).Sorted = True
  frmNodeInfo.TV.Nodes(TV.SelectedItem.Key).Expanded = True
  
  'If this node is a parent
  If ParentFiles.IsParent(TV.SelectedItem.Key) Then
    'If there are any children, add each one
    If ParentFiles.Item(TV.SelectedItem.Key).Count > 0 Then
      For x = 1 To ParentFiles.Item(TV.SelectedItem.Key).Count
        frmNodeInfo.TV.Nodes.Add TV.SelectedItem.Key, tvwChild, ParentFiles.Item(TV.SelectedItem.Key).Item(x).Path, TV.Nodes(ParentFiles.Item(TV.SelectedItem.Key).Item(x).Path).Text, TV.Nodes(ParentFiles.Item(TV.SelectedItem.Key).Item(x).Path).Image
      Next x
    End If
  End If

End Function

'Show the violations in the listbox
Public Function ShowNodeInfo_Violations(TV As TreeView, LV As ListView)

  Dim TotalViolations As Integer
  Dim LowVCount As Integer
  Dim NormalVCount As Integer
  Dim HighVCount As Integer
  Dim lFile As typFile
  Dim x As Integer

  TotalViolations = 0
  LowVCount = 0
  NormalVCount = 0
  HighVCount = 0
  
  'Basically just cycles through the listview and checks to see if the name of the
  'subject node is present anywhere
  If Len(TV.SelectedItem.Image) >= 5 Then
    If Left$(TV.SelectedItem.Image, 5) <> "CHECK" And Left$(TV.SelectedItem.Image, 5) <> "CROSS" Then
      TotalViolations = -1
    End If
  Else
    TotalViolations = -1
  End If
  
  If LV.ListItems.Count > 0 And TotalViolations = 0 Then
    lFile = BuildFileType(TV.SelectedItem.Key)
    For x = 1 To LV.ListItems.Count
      If LV.ListItems(x).Text = lFile.Name & "." & lFile.Ext & " - " & lFile.FullName Then
        Select Case LV.ListItems(x).ListSubItems(2).Text
          Case "Low"
            LowVCount = LowVCount + 1
          Case "Normal"
            NormalVCount = NormalVCount + 1
          Case "High"
            HighVCount = HighVCount + 1
        End Select
      End If
    Next x
  End If
    
  'Displays the discovered data
  TotalViolations = TotalViolations + LowVCount + NormalVCount + HighVCount
  
  If TotalViolations = -1 Then
    frmNodeInfo.LV.ListItems(1).ListSubItems(1).Text = "Node Not Scaned"
    frmNodeInfo.LV.ListItems(2).ListSubItems(1).Text = ""
    frmNodeInfo.LV.ListItems(3).ListSubItems(1).Text = ""
    frmNodeInfo.LV.ListItems(4).ListSubItems(1).Text = ""
  Else
    frmNodeInfo.LV.ListItems(1).ListSubItems(1).Text = LowVCount
    frmNodeInfo.LV.ListItems(2).ListSubItems(1).Text = NormalVCount
    frmNodeInfo.LV.ListItems(3).ListSubItems(1).Text = HighVCount
    frmNodeInfo.LV.ListItems(4).ListSubItems(1).Text = TotalViolations
  End If

End Function

'Searches a given key's children for violations
Public Function SearchKeysChildren(tTopKey As String, TV As TreeView, LV As ListView, tProgressBarContainer As PictureBox, tProgressBar As PictureBox, tProgressBarOffset As Integer, tProgressCaption As Label, SD As clsDictionary)

  Dim tKey As String
  Dim tLastKey As String
  
  'If the key has no children, get out
  If TV.Nodes(tTopKey).Children < 1 Then Exit Function
  
  'Set the last sibling key
  tLastKey = TV.Nodes(tTopKey).Child.LastSibling.Key
  
  'Set the first sibling key and search it - and all of it's children
  tKey = TV.Nodes(tTopKey).Child.FirstSibling.Key
  SearchFile tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, TV.Nodes(tKey)
  SearchKeysChildren tKey, TV, LV, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD
  
  Do Until tKey = tLastKey
    'Set the next sibling key and search it - and all of it's children
    tKey = TV.Nodes(tKey).Next.Key
    SearchFile tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD, LV, TV.Nodes(tKey)
    SearchKeysChildren tKey, TV, LV, tProgressBarContainer, tProgressBar, tProgressBarOffset, tProgressCaption, SD
  Loop

End Function

'Basically just copies all of the registry values from ...\Source Search 1.0\ to ...\Source Search\
Private Function SetupWizard_UpgradeOptions()

  Dim x As Integer
  Dim d As String
  Dim tCount As Integer
  Dim tPhrase As String
  Dim tSeverity As Integer
  Dim tDescription As String
  
  d = GetSetting("Source Search 1.0", "Dictionary", "Total Definitions", 0)
  SaveSetting AppTitle, "Dictionary", "Total Definitions", d
  If IsNumeric(d) Then tCount = Int(d) Else Exit Function
    
  If tCount <= 0 Then Exit Function
  
  For x = 1 To tCount
  
    tPhrase = GetSetting("Source Search 1.0", "Dictionary", "Def" & x & "_Phrase", "")
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Phrase", tPhrase
    If tPhrase = "" Then Exit Function
    
    d = GetSetting("Source Search 1.0", "Dictionary", "Def" & x & "_Severity", 1)
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Severity", d
    If IsNumeric(d) Then tSeverity = Int(d) Else Exit Function
    If tSeverity > 2 Then tSeverity = 2
    If tSeverity < 0 Then tSeverity = 0
    
    tDescription = GetSetting("Source Search 1.0", "Dictionary", "Def" & x & "_Description", "")
    SaveSetting AppTitle, "Dictionary", "Def" & x & "_Description", tDescription
            
  Next x
  
  FlySearch = GetSetting("Source Search 1.0", "Options", "FlySearch", True)
  SaveDictionary = Int(GetSetting("Source Search 1.0", "Options", "SaveDictionary", 1))
  IgnoreComments = GetSetting("Source Search 1.0", "Options", "IgnoreComments", True)
  RemoveZips = GetSetting("Source Search 1.0", "Options", "RemoveZips", True)
  
  SaveSetting AppTitle, "Options", "FlySearch", FlySearch
  SaveSetting AppTitle, "Options", "SaveDictionary", SaveDictionary
  SaveSetting AppTitle, "Options", "IgnoreComments", IgnoreComments
  SaveSetting AppTitle, "Options", "RemoveZips", RemoveZips
  
End Function



'
