Attribute VB_Name = "VirtualTubMac"
Option Explicit

'Globals to ensure menu ID's, button ID's etc. increment correctly
Public VTubFile As String
Public VTubFileArray() As String
Public VTubCurrentFileNumber As Long
Public VTubMaxDepthExceeded As Boolean
Public VTubFileCount As Long
Public VTubLastModified As String

Sub GetVTubContent(Optional FromScratch As Boolean)
'Get content for dynamic VTub menu

    Dim VTubPath As String
    Dim VTubPOSIX As String
    Dim InputFile
    
    Dim VTubRootNode As CommandBarControl
    Dim c
    Dim FileElement
    Dim LineSplit
    Dim i
    Dim j
    
    'Check if VTubXMLFile exists
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & VTubPath & "VTub.txt" & """" & Chr(13) & "end tell") = "false" Then
        'If no VTub file, add a Create button
        Set Button = VTubRootNode.Controls.Add(Type:=msoControlButton)
        Button.Caption = "Create VTub"
        Button.Tag = "CreateVTub"
        Button.FaceId = 1394 'Partial box
        Button.OnAction = "VirtualTub.VTubCreate"
        Set Button = Nothing
        Exit Sub
    End If

    'If VTubRefreshPrompt is turned on, check if Tub is out of date by comparing date modified of files to VTub file
    If GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True) = True Then
        VTubMaxDepthExceeded = False
        VTubLastModified = ""
        Call VirtualTub.VTubFileCounter(VTubPath)
        If MacScript("do shell script ""stat -f '%m %N' " & VTubPOSIX & "VTub.txt" & "| cut -d ' ' -f 1""") <> VTubLastModified Then
            If MsgBox("The VTub has not been refreshed since you last changed files. Refresh Now?", vbYesNo) = vbYes Then
                Call VirtualTub.VTubRefresh
                Exit Sub
            End If
        End If
    End If

    'Open and read the VTub file
    InputFile = FreeFile
    Open VTubPath & "VTub.txt" For Input As #InputFile
    VTubFile = Input$(LOF(InputFile), InputFile)
    Close #InputFile

    'Split the VTub file into an array
    VTubFileArray = Split(VTubFile, "!#!FILE END!#!" & Chr(13))

    'Loop each file in the VTub
    For i = 0 To UBound(VTubFileArray) - 1
    
        'Split the file info into lines
        FileElement = Split(VTubFileArray(i), Chr(13))
        
        Set Menu = VTubRootNode.Controls.Add(Type:=msoControlPopup)
        Menu.Caption = Replace(Replace(FileElement(0), VTubPath, ""), ":", "/")
        Menu.Tag = FileElement(0)
    
        'Starting with first bookmark, loop and create buttons
        For j = 2 To UBound(FileElement) - 1
            LineSplit = Split(FileElement(j), "!#!")
            Set Button = Menu.Controls.Add(Type:=msoControlButton)
            Select Case Left(LineSplit(1), 3) 'Get left 3 characters of Bookmark name
                Case "Poc" 'Pocket
                    Button.Caption = LineSplit(2)
                Case "Hat" 'Hat
                    Button.Caption = vbTab & LineSplit(2)
                Case "Blo" 'Block
                    Button.Caption = vbTab & vbTab & LineSplit(2)
                Case Else
                    Button.Caption = LineSplit(2)
            End Select
            
            Button.Tag = LineSplit(0) & "!#!" & LineSplit(1)
            Button.OnAction = "VirtualTub.VTubInsertBookmark"
        Next
    Next

    'Default buttons
    Set Button = VTubRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = ""
    Button.Tag = "VTubSeparator1"
    Button.Enabled = False
    
    Set Button = VTubRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "Refresh VTub"
    Button.Tag = "RefreshVTub"
    Button.OnAction = "Toolbar.AssignButtonActions"
    Button.FaceId = 8085 'Blue Refresh
    
    Set Button = VTubRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "Recreate VTub"
    Button.Tag = "RecreateVTub"
    Button.OnAction = "VirtualTub.VTubCreate"
    Button.FaceId = 1399 'Empty Box

    Set Button = VTubRootNode.Controls.Add(Type:=msoControlButton)
    Button.Caption = "VTub Settings"
    Button.Tag = "VTubSettings"
    Button.OnAction = "Settings.ShowSettingsForm"
    Button.FaceId = 2144 'Gears

    'Set template as saved to avoid prompts
    ActiveDocument.AttachedTemplate.Saved = True
    
    'Clean up
    Set Menu = Nothing
    Set Button = Nothing

    Exit Sub

Handler:
    Set Menu = Nothing
    Set Button = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub VTubCreate()

    Dim VTubPath As String
    Dim VTubFilePath As String
    Dim Subfolders As Variant
    Dim Subfolder

    Dim OutputFile
    
    Dim VTubRootNode As CommandBarControl
    Dim c
    
    On Error GoTo Handler

    'Get VTubPath from settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "?")
    If Right(VTubPath, 1) <> ":" Then VTubPath = VTubPath & ":" 'Append trailing : if missing

    'Check if VTub already exists - prompt to refresh instead or delete file
    If MacScript("tell application ""Finder""" & Chr(13) & "exists file """ & VTubPath & "VTub.txt" & """" & Chr(13) & "end tell") = "true" Then
        If MsgBox("VTub already exists - you can update it with the ""Refresh"" button in the VTub menu. Recreate from scratch instead?", vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
        Else
            Filesystem.KillFileOnMac VTubPath & "VTub.txt"
        End If
    End If

    'Reset counters and get initial count of files in the VTub
    VTubFileCount = 0
    VTubMaxDepthExceeded = False
    Call VirtualTub.VTubFileCounter(VTubPath)

    'Check if large number of files
    If VTubFileCount > 20 Then
        If MsgBox("You have a large number of files (>20) in the VTub. This could take a few minutes - okay?", vbYesNo, "You sure?") = vbNo Then Exit Sub
    End If

    'Warn if more than 2 folder levels
    If VTubMaxDepthExceeded = True Then MsgBox "VTub can only handle one level of subfolders - files deeper than one subfolder will be ignored.", vbOKOnly

    'Initialize file array
    Erase VTubFileArray

    'Reset counters
    VTubCurrentFileNumber = 0

    'Warn about files opening and closing
    MsgBox "While the VTub is being created, you will see files being opened and closed automatically. Be patient - this could take a few minutes, and you'll be notified when the process is complete.", vbOKOnly

    'Show progress bar
    ProgressBar = "Creating VTub - File 0 of " & VTubFileCount
    Application.StatusBar = ProgressBar

    'Process each subfolder first
    Subfolders = Split(Filesystem.GetSubfoldersInFolder(VTubPath), Chr(10))
    For Each Subfolder In Subfolders
        Call VirtualTub.VTubProcessFolder(Subfolder)
    Next Subfolder
    
    'Process main folder
    Call VirtualTub.VTubProcessFolder(VTubPath)

    'Update progress form as complete
    ProgressBar = "VTub Creation Complete"
    Application.StatusBar = ProgressBar
    
    'Create VTub file from the array
    VTubFile = Join(VTubFileArray, "")
    
    'Exit if nothing in VTub
    If Len(VTubFile) = 0 Then
        MsgBox "VTub creation failed!"
        Exit Sub
    End If
    
    'Save file
    VTubFilePath = VTubPath & "VTub.txt"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, VTubFile
    Close #OutputFile

    'Clear existing VTub Menu
    Set VTubRootNode = CommandBars.FindControl(Tag:="VirtualTub")
    For Each c In VTubRootNode.Controls
        c.Delete
    Next c
    
    'Notify
    MsgBox "VTub successfully created!"

    'Clean Up
    Erase VTubFileArray
    Set VTubRootNode = Nothing

    Exit Sub

Handler:
    Erase VTubFileArray
    Set VTubRootNode = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description
End Sub

Private Sub VTubProcessFolder(Folder)
   
    Dim FileList As Variant
    Dim i
    Dim j
    
    Dim FilePOSIX As String
    Dim OldDateModified As String
    Dim NewDateModified As String

    Dim Refresh As Boolean
    Dim FileBookmarks As String

    'Turn on error-checking
    On Error GoTo Handler

    'Get files in current folder
    FileList = Split(Filesystem.GetFilesInFolder(Folder), Chr(10))

    'Iterate through each file in the folder
    For i = 0 To UBound(FileList)
        
            'Increment the Progress Bar
            VTubCurrentFileNumber = VTubCurrentFileNumber + 1
            ProgressBar = Str(Round((VTubCurrentFileNumber / VTubFileCount) * 100, 0)) & "% - " & "Processing File " & FileList(i) & " (" & VTubCurrentFileNumber & " of " & VTubFileCount & ")"
            Application.StatusBar = ProgressBar

            'Loop through files in the VTubArray - if there's a match, we're refreshing a file
            Refresh = False
            If Not Not VTubFileArray Then
                For j = 0 To UBound(VTubFileArray)
                    If InStr(VTubFileArray(j), FileList(i)) Then 'If VTub Element includes the filename, there's a match
                        Refresh = True
                        FilePOSIX = MacScript("get quoted form of POSIX path of """ & Trim(FileList(i)) & """")
                        OldDateModified = Split(VTubFileArray(j), Chr(13))(1)
                        NewDateModified = MacScript("do shell script ""stat -f '%m %N' " & FilePOSIX & "| cut -d ' ' -f 1""")
                        If OldDateModified <> NewDateModified Then 'Only update if the file has been modified
                            FileBookmarks = VTubProcessFile(FileList(i)) 'Reprocess file
                            
                            'Update VTub with new file info
                            NewDateModified = MacScript("do shell script ""stat -f '%m %N' " & FilePOSIX & "| cut -d ' ' -f 1""")
                            VTubFileArray(j) = FileList(i) & Chr(13) & NewDateModified & Chr(13) & FileBookmarks & Chr(13) & "!#!FILE END!#!" & Chr(13)
                            
                        End If
                    End If
                Next
            End If

            'No match found, we're creating file node from scratch
            If Refresh = False Then
                'Increase size of array by one
                If Not Not VTubFileArray Then
                    ReDim Preserve VTubFileArray(0 To UBound(VTubFileArray) + 1) As String
                Else
                    ReDim Preserve VTubFileArray(0) As String
                End If
                
                'Process file
                FileBookmarks = VTubProcessFile(FileList(i))
                
                'Add file as last element
                NewDateModified = MacScript("do shell script ""stat -f '%m %N' " & FilePOSIX & "| cut -d ' ' -f 1""")
                VTubFileArray(UBound(VTubFileArray)) = Trim(FileList(i)) & Chr(13) & NewDateModified & Chr(13) & FileBookmarks & "!#!FILE END!#!" & Chr(13)
                
            End If
    Next
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Function VTubProcessFile(FileName) As String

    Dim FileBookmarks As String

    Dim PocketBMRange As Range
    Dim HatBMRange As Range
    Dim BlockBMRange As Range

    Dim PocketBMName As String
    Dim HatBMName As String
    Dim BlockBMName As String

    Dim PocketOpen As Boolean
    Dim HatOpen As Boolean
    Dim BlockOpen As Boolean

    Dim p As Paragraph
    Dim pCount As Long
    Dim pFix As String

    On Error GoTo Handler

    'Open the file and activate it
    Documents.Open FileName:=FileName
    Documents(FileName).Activate

    'Delete all bookmarks
    Call VirtualTub.RemoveBookmarks

    'Initialize bookmark ranges
    Set PocketBMRange = ActiveDocument.Range
    Set HatBMRange = ActiveDocument.Range
    Set BlockBMRange = ActiveDocument.Range

    'Open a new bookmark if first paragraph is a heading
    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel1 Then
        PocketOpen = True
    Else
        PocketOpen = False
    End If

    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel2 Then
        HatOpen = True
    Else
        HatOpen = False
    End If

    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel3 Then
        BlockOpen = True
    Else
        BlockOpen = False
    End If

    'Initialize bookmark names
    pCount = 0
    PocketBMName = "PocketBM1"
    HatBMName = "HatBM1"
    BlockBMName = "BlockBM1"

    'Loop all paragraphs
    For Each p In Documents(FileName).Paragraphs
        pCount = pCount + 1

        'If end of file, close all bookmarks
        If p.Range.End = Documents(FileName).Range.End Then
            PocketBMRange.End = Documents(FileName).Range.End
            HatBMRange.End = Documents(FileName).Range.End
            BlockBMRange.End = Documents(FileName).Range.End
            If PocketOpen = True Then Documents(FileName).Bookmarks.Add PocketBMName, PocketBMRange
            If HatOpen = True Then Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
            If BlockOpen = True Then Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
        End If

        'Process depending on outline level
        Select Case p.outlineLevel

        Case Is = 1 'Pocket

            If Len(p.Range.Text) > 0 Then

                'Close open bookmarks
                If PocketOpen = True Then
                    PocketBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add PocketBMName, PocketBMRange
                    PocketOpen = False
                End If
                If HatOpen = True Then
                    HatBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
                    HatOpen = False
                End If
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If

                'Start a new bookmark
                PocketBMRange.Start = p.Range.Start
                PocketBMName = "PocketBM" & pCount
                PocketOpen = True

                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000) 'Limit length to 1000 characters to avoid breaking XML
                If pFix = "" Then pFix = "-"

                'Append
                FileBookmarks = FileBookmarks & FileName & "!#!" & PocketBMName & "!#!" & pFix & Chr(13)
                
            End If

        Case Is = 2 'Hat
            If Len(p.Range.Text) > 0 Then

                'Close open bookmarks
                If HatOpen = True Then
                    HatBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
                    HatOpen = False
                End If
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If

                'Start a new bookmark
                HatBMRange.Start = p.Range.Start
                HatBMName = "HatBM" & pCount
                HatOpen = True

                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000)
                If pFix = "" Then pFix = "-"

                'Append
                FileBookmarks = FileBookmarks & FileName & "!#!" & HatBMName & "!#!" & pFix & Chr(13)
                
            End If

        Case Is = 3 'Block
            If Len(p.Range.Text) > 0 Then

                'Close open bookmarks
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If

                'Start a new bookmark
                BlockBMRange.Start = p.Range.Start
                BlockBMName = "BlockBM" & pCount
                BlockOpen = True

                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000)
                If pFix = "" Then pFix = "-"

                'Append
                FileBookmarks = FileBookmarks & FileName & "!#!" & BlockBMName & "!#!" & pFix & Chr(13)
                
            End If

        Case Else
            'Do nothing

        End Select

    Next p

    'Close file and save changes
    Documents(FileName).Close SaveChanges:=wdSaveChanges

    'Return the updated file node
    VTubProcessFile = FileBookmarks

    Exit Function

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Function

Sub VTubFileCounter(Folder As String)

    Dim POSIXFolder As String

    On Error Resume Next

    'Convert to POSIX path
    POSIXFolder = MacScript("get quoted form of POSIX path of """ & Folder & """")

    'If files found deeper than 2 levels, set the warning flag
    If MacScript("do shell script ""find " & POSIXFolder & " -name '*.doc*' -mindepth 3""") <> "" Then
        VTubMaxDepthExceeded = True
    Else
        VTubMaxDepthExceeded = False
    End If
    
    'Get count of all .doc or .docx files in first two levels
    VTubFileCount = MacScript("do shell script ""find " & POSIXFolder & " -name '*.doc*' -maxdepth 2 | wc -l""")
    
    'Save the most recent date modified
    VTubLastModified = MacScript("do shell script ""find " & POSIXFolder & " -type f -print0 | xargs -0 stat -f '%m %N' | sort -rn | head -1 | cut -d ' ' -f 1""")
    
    Exit Sub

Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub VTubRefresh()

    Dim VTubPath As String

    Dim Subfolders As Variant
    Dim Subfolder
    
    Dim InputFile
    Dim OutputFile
    
    Dim VTubRootNode As CommandBarControl
    Dim c
    
    On Error GoTo Handler

    'Verify before proceeding
    If MsgBox("Are you sure you want to refresh the VTub?", vbOKCancel) = vbCancel Then Exit Sub

    'Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "?")
    If Right(VTubPath, 1) <> ":" Then VTubPath = VTubPath & ":" 'Append trailing : if missing

    'Open and read the VTub file
    InputFile = FreeFile
    Open VTubPath & "VTub.txt" For Input As #InputFile
    VTubFile = Input$(LOF(InputFile), InputFile)
    Close #InputFile

    'Split the VTub file into an array
    VTubFileArray = Split(VTubFile, "!#!END FILE!#!" & Chr(13))

    'Reset counters and get initial count of files in the VTub
    VTubFileCount = 0
    Call VirtualTub.VTubFileCounter(VTubPath)

    'Reset counters
    VTubCurrentFileNumber = 0

    'Show progress bar
    ProgressBar = "Refreshing VTub - File 0 of " & VTubFileCount
    Application.StatusBar = ProgressBar

    'Process each subfolder first
    Subfolders = Split(Filesystem.GetSubfoldersInFolder(VTubPath), Chr(13))
    For Each Subfolder In Subfolders
        Call VirtualTub.VTubProcessFolder(Subfolder)
    Next Subfolder
    
    'Process main folder
    Call VirtualTub.VTubProcessFolder(VTubPath)

    'Update progress form as complete
    ProgressBar = "VTub Refresh Complete!"
    Application.StatusBar = ProgressBar
    
    'Create VTub file from the array and save
    VTubFile = Join(VTubFileArray, "")
    
    OutputFile = FreeFile
    Open VTubPath & "VTub.txt" For Output As #OutputFile
    Print #OutputFile, VTubFile
    Close #OutputFile

    'Clear existing VTub Menu
    Set VTubRootNode = CommandBars.FindControl(Tag:="VirtualTub")
    For Each c In VTubRootNode.Controls
        c.Delete
    Next c
    
    'Notify
    MsgBox "VTub successfully refreshed!"

    'Clean Up
    Set VTubRootNode = Nothing

    Exit Sub

Handler:
    Set VTubRootNode = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub RemoveBookmarks()

    Dim b As Bookmark
    For Each b In ActiveDocument.Bookmarks
        b.Delete
    Next b

End Sub

Sub VTubInsertBookmark()

    Dim PressedControl As CommandBarButton
    Set PressedControl = CommandBars.ActionControl
    
    'Insert bookmark - get the file path and bookmark name by splitting the tag attribute on the !#! delimiter
    Selection.InsertFile Split(PressedControl.Tag, "!#!", 2)(0), Split(PressedControl.Tag, "!#!", 2)(1)

    Set PressedControl = Nothing

End Sub



