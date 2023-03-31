Attribute VB_Name = "VirtualTub"
'@IgnoreModule IndexedUnboundDefaultMemberAccess, ProcedureNotUsed
Option Explicit

'@Ignore ParameterNotUsed, ProcedureCanBeWrittenAsFunction
Public Sub GetVTubContent(ByVal c As IRibbonControl, ByRef returnedVal As Variant)
' Get content for dynamic menu from JSON file
    Dim VTubPath As String
    Dim Folder As clsFolder
    Dim SubFolder As clsFolder
    Dim File As clsFile
    Dim MostRecent As Date
    Dim i As Long
    Dim j As Long
       
    On Error Resume Next
        
    ' Get VTubPath from Settings and make sure it exists
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    If VTubPath = "" Or VTubPath = "?" Then
        If MsgBox("You haven't configured a VTub location in the Verbatim settings. Open Settings?", vbYesNo) = vbYes Then
            UI.ShowForm "Settings"
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' Append trailing separator if missing
    If Right$(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Check if VTub file exists
    If Filesystem.FileExists(VTubPath & "VTub.json") = False Then
        ' If no file, return a button to create it
        returnedVal = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
        returnedVal = returnedVal & "<button id=""CreateVTub"" label=""Create VTub"" onAction=""VirtualTub.VTubCreateButton"" image=""tub""" & " />"
        returnedVal = returnedVal & "</menu>"
        Exit Sub
    End If

    ' If VTubRefreshPrompt is turned on, check if Tub is out of date by comparing date modified of files to JSON file
    ' Have to loop all files becuase Word changes folder modified date when opening docs
    If GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True) = True Then
        Set Folder = Filesystem.GetFolder(VTubPath)
        MostRecent = Folder.DateLastModified

        For i = 1 To Folder.Subfolders.Count
            Set SubFolder = Filesystem.GetFolder(Folder.Subfolders.Item(i))
            For j = 1 To SubFolder.Files.Count
                If Right$(SubFolder.Files.Item(j), 4) = "docx" And Left$(SubFolder.Files.Item(j), 1) <> "~" Then
                    Set File = Filesystem.GetFile(SubFolder.Files.Item(j))
                    If File.DateLastModified > MostRecent Then MostRecent = File.DateLastModified
                End If
            Next j
        Next
        
        For i = 1 To Folder.Files.Count
            If Right$(Folder.Files.Item(i), 4) = "docx" And Left$(Folder.Files.Item(i), 1) <> "~" Then
                Set File = Filesystem.GetFile(Folder.Files.Item(i))
                If File.DateLastModified > MostRecent Then MostRecent = File.DateLastModified
            End If
        Next i
        
        If Filesystem.GetFile(VTubPath & "VTub.json").DateLastModified < MostRecent Then
            If MsgBox("The VTub has not been refreshed since you last changed files. Refresh Now?", vbYesNo) = vbYes Then
                VirtualTub.VTubRefresh
                Exit Sub
            End If
        End If
    End If

    returnedVal = VirtualTub.VTubConvertToXML
    
    On Error GoTo 0
    
    Exit Sub
End Sub

Private Function VTubConvertToXML() As String
    Dim VTubPath As String
    Dim JSON As String
    Dim RootMenu As Object
    Dim key As Variant
    Dim Menu As Dictionary
    Dim xml As String
    
    On Error GoTo Handler
    
    ' Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    If Right$(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Load and parse JSON from file
    JSON = Filesystem.ReadFile(VTubPath & "VTub.json")
    Set RootMenu = JSONTools.ParseJson(JSON)

    ' Start the XML menu
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
    
    ' Convert each subkey Dictionary (file/folder) in the JSON into the XML equivalent
    ' Not truly recursive because it can only go to 5 menu levels without breaking the ribbon menu
    For Each key In RootMenu.Keys
        If key <> "FileCount" And key <> "FolderCount" And key <> "DateLastModified" Then
            Set Menu = RootMenu(key)
            xml = xml & VirtualTub.ConvertDictionaryToXML(Menu)
        End If
    Next key
  
    ' Add default buttons
    xml = xml & "<menuSeparator id=""VTubSeparator"" />"
    #If Mac Then
        xml = xml & "<button id=""RefreshVTub"" label=""Refresh VTub"" onAction=""VirtualTub.VTubRefreshButton"" imageMso=""Refresh"" />"
    #Else
        xml = xml & "<button id=""RefreshVTub"" label=""Refresh VTub"" onAction=""VirtualTub.VTubRefreshButton"" imageMso=""AccessRefreshAllLists"" />"
    #End If
    xml = xml & "<button id=""RecreateVTub"" label=""Recreate VTub"" onAction=""VirtualTub.VTubCreateButton"" image=""tub"" />"
    xml = xml & "<button id=""VTubSettings"" label=""VTub Settings"" onAction=""VirtualTub.VTubSettingsButton"" imageMso=""_3DLightingFlatClassic"" />"
    
    xml = xml & "</menu>"
    
    VTubConvertToXML = xml
    
    Exit Function
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Private Function ConvertDictionaryToXML(ByVal d As Dictionary) As String
    Dim xml As String
    Dim MenuID As String
    Dim Children As Dictionary
    Dim Child As Dictionary
    Dim key As Variant
    
    ' Use a random menu ID to avoid collisions
    MenuID = Format$(Int(Rnd * 10 ^ 8), "00000000")
     
    ' Children means we have to recurse to the bottom of the tree
    If d.Exists("Children") Then
        Set Children = d.Item("Children")
        
        ' Headings in files are a <splitbutton> to insert the whole heading or any children
        If d.Item("MenuType") = "Heading" Then
            xml = xml & "<splitButton "
            xml = xml & "id=""VTub" & MenuID & """>" & vbCrLf
       
            MenuID = MenuID & "1"
            xml = xml & "<button "
            xml = xml & "id=""VTub" & MenuID & """ "
            xml = xml & "label=""" & d.Item("Label") & """ "
            xml = xml & "tag=""" & d.Item("Path") & "!#!" & d.Item("Name") & """ "
            xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
            #If Mac Then
                xml = xml & "imageMso=""TextDirection"" "
            #Else
                xml = xml & "imageMso=""ExportTextFile"" "
            #End If
            xml = xml & "/>" & vbCrLf
           
            MenuID = MenuID & "1"
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & MenuID & """ "
            xml = xml & ">" & vbCrLf
        
        ' File or Folder node is a <menu>
        Else
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & MenuID & """ "
            xml = xml & "label=""" & d.Item("Name") & """ "
            If d.Item("MenuType") = "Folder" Then
                xml = xml & "imageMso=""Folder"" "
            Else
                xml = xml & "imageMso=""FileSaveAsWordDocx"" "
            End If
            xml = xml & "tag=""" & d.Item("Path") & "!#!" & d.Item("DateLastModified") & """"
            xml = xml & ">" & vbCrLf
        End If
        
        ' Recurse any sub-menus in the dictionary
        For Each key In Children.Keys
            Set Child = Children.Item(key)
            ' Sub-children need to keep recursing
            If Child.Exists("Children") Then
                xml = xml & ConvertDictionaryToXML(Child)
            
            ' Add a leaf node button to the menu
            Else
                MenuID = MenuID & "1"
                xml = xml & "<button "
                xml = xml & "id=""VTub" & MenuID & """ "
                xml = xml & "label=""" & Child.Item("Label") & """ "
                xml = xml & "tag=""" & Child.Item("Path") & "!#!" & Child.Item("Name") & """ "
                xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
                #If Mac Then
                    xml = xml & "imageMso=""TextDirection"" "
                #Else
                    xml = xml & "imageMso=""ExportTextFile"" "
                #End If
                xml = xml & "/>" & vbCrLf
            End If
        Next key
        
        ' Close the menu/splitButton
        xml = xml & "</menu>" & vbCrLf
        If d.Item("MenuType") = "Heading" Then
            xml = xml & "</splitButton>" & vbCrLf
        End If

    ' No children means a terminal leaf node (e.g. a terminal heading or empty file), so just add a buttion or empty menu
    Else
        If d.Item("MenuType") = "Heading" Then
            xml = xml & "<button "
            xml = xml & "id=""VTub" & MenuID & """ "
            xml = xml & "label=""" & d.Item("Label") & """ "
            xml = xml & "tag=""" & d.Item("Path") & "!#!" & d.Item("Name") & """ "
            xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
            #If Mac Then
                xml = xml & "imageMso=""TextDirection"" "
            #Else
                xml = xml & "imageMso=""ExportTextFile"" "
            #End If
            xml = xml & "/>" & vbCrLf
        Else
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & MenuID & """ "
            xml = xml & "label=""" & d.Item("Name") & """ "
            If d.Item("MenuType") = "Folder" Then
                xml = xml & "imageMso=""Folder"" "
            Else
                xml = xml & "imageMso=""FileSaveAsWordDocx"" "
            End If
            xml = xml & "tag=""" & d.Item("Path") & "!#!" & d.Item("DateLastModified") & """"
            xml = xml & ">" & vbCrLf
            xml = xml & "</menu>" & vbCrLf
        End If
    End If
    
    ConvertDictionaryToXML = xml
End Function

'@Ignore ParameterNotUsed
Public Sub VTubRefreshButton(ByVal c As IRibbonControl)
    If MsgBox("Are you sure you want to refresh the VTub?", vbOKCancel) = vbCancel Then Exit Sub
    VirtualTub.VTubRefresh
End Sub

'@Ignore ParameterNotUsed
Public Sub VTubCreateButton(ByVal c As IRibbonControl)
    If MsgBox("Are you sure you want to create the VTub from scratch?", vbYesNo, "Create VTub?") = vbNo Then Exit Sub
    VirtualTub.VTubCreate
End Sub

'@Ignore ParameterNotUsed
Public Sub VTubSettingsButton(ByVal c As IRibbonControl)
    UI.ShowForm "Settings"
End Sub

Public Sub VTubInsertBookmark(ByVal c As IRibbonControl)
    ' Insert bookmark - get the file path and bookmark name by splitting the tag attribute on the !#! delimiter
    Selection.InsertFile Split(c.Tag, "!#!", 2)(0), Split(c.Tag, "!#!", 2)(1)
End Sub

Public Sub VTubCreate()
    Dim VTubPath As String
    Dim Folder As clsFolder
    Dim SubFolder As clsFolder
    Dim File As clsFile
    
    Dim RootMenu As Dictionary
    Dim SubfolderMenu As Dictionary
    Dim FileMenu As Dictionary
    Dim Headings As Dictionary
    Dim Children As Dictionary
        
    Dim FileCount As Long
    FileCount = 0
    
    Dim DepthExceeded As Boolean
    
    Dim i As Long
    Dim j As Long
    Dim CurrentFileCount As Long
    CurrentFileCount = 0
    Dim ProgressPct As Double
    Dim ProgressForm As frmProgress
    
    Dim JSON As String
    Dim OutputFile As Variant
    
    On Error GoTo Handler
    
    ' Initialize the root JSON object
    Set RootMenu = New Dictionary
    
    ' Get VTubPath from settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    
    If VTubPath = "" Or VTubPath = "?" Then
        If MsgBox("You haven't configured a folder for the VTub. Open Settings?", vbYesNo, "Open Settings?") = vbYes Then
            UI.ShowForm "Settings"
        End If
        Exit Sub
    End If
    
    ' Append trailing \ if missing
    If Right$(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    #If Mac Then
        ' Request permissions for all files in VTub at once so the user doesn't have to allow access to each file individually
        Filesystem.RequestFolderAccess VTubPath
    #End If
    
    Set Folder = Filesystem.GetFolder(VTubPath)
      
    ' Check the VTub depth and file count
    For i = 1 To Folder.Subfolders.Count
        Set SubFolder = Filesystem.GetFolder(Folder.Subfolders.Item(i))
        If SubFolder.Subfolders.Count > 0 Then DepthExceeded = True
        FileCount = FileCount + SubFolder.Files.Count
    Next
    
    FileCount = FileCount + Folder.Files.Count
           
    If FileCount > 20 Then
        If MsgBox("You have a large number of files (>20) in the VTub. This could take a few minutes - okay?", vbYesNo, "You sure?") = vbNo Then Exit Sub
    End If
    
    If DepthExceeded = True Then MsgBox "VTub can only handle one level of subfolders - files deeper than one subfolder will be ignored.", vbOKOnly

    ' Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Creating VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of approximately " & FileCount
    ProgressForm.Show

    ' Count numbers are approximate because they include non-docx and temp files
    RootMenu.Add "FileCount", FileCount
    RootMenu.Add "FolderCount", Folder.Subfolders.Count

   ' Iterate through each subfolder in first depth level - XML menu is limited to 5 levels and we need 3 for the file contents
    For i = 1 To Folder.Subfolders.Count
        Set SubFolder = Filesystem.GetFolder(Folder.Subfolders.Item(i))
        Set SubfolderMenu = New Dictionary
        SubfolderMenu.Add "MenuType", "Folder"
        SubfolderMenu.Add "Name", SubFolder.Name
        SubfolderMenu.Add "Path", SubFolder.Path
        
        Set Children = New Dictionary
        
        For j = 1 To SubFolder.Files.Count
            ' Trap for cancel button on Progress Form
            If ProgressForm.Visible = False Then Exit Sub
            
            If Right$(SubFolder.Files.Item(j), 4) = "docx" And Left$(SubFolder.Files.Item(j), 1) <> "~" Then
                ' Increment the progress form
                CurrentFileCount = CurrentFileCount + 1
                ProgressPct = CurrentFileCount / FileCount
                ProgressForm.lblCaption.Caption = Str$(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & CurrentFileCount & " of " & FileCount
                ProgressForm.lblFile.Caption = "Processing " & SubFolder.Files.Item(j)
                ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
                If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
                
                DoEvents ' Necessary for Progress form to update
                           
                ' Process each file
                Set File = Filesystem.GetFile(SubFolder.Files.Item(j))
                Set FileMenu = New Dictionary
                FileMenu.Add "MenuType", "File"
                FileMenu.Add "Name", File.Name
                FileMenu.Add "Path", File.Path
                
                ' Convert headings to bookmarks
                Set Headings = VirtualTub.VTubProcessFile(File.Path)
                FileMenu.Add "Children", Headings
                
                ' Re-initialize the file to get the new modified timestamp
                Set File = Filesystem.GetFile(SubFolder.Files.Item(j))
                FileMenu.Add "DateLastModified", Format$(File.DateLastModified)
                
                ' Have to double escape \'s to not blow up the JSON parser
                Children.Add Replace(File.Path, "\", "\\"), FileMenu
            End If
        Next
        
        SubfolderMenu.Add "Children", Children
        SubfolderMenu.Add "DateLastModified", Format$(SubFolder.DateLastModified)

        RootMenu.Add Replace(SubFolder.Path, "\", "\\"), SubfolderMenu
    Next
    
    ' Process top-level files
    For i = 1 To Folder.Files.Count
        ' Trap for cancel button on Progress Form
        If ProgressForm.Visible = False Then Exit Sub
        
        If Right$(Folder.Files.Item(i), 4) = "docx" And Left$(Folder.Files.Item(i), 1) <> "~" Then
        
            ' Increment the progress form
            CurrentFileCount = CurrentFileCount + 1
            ProgressPct = CurrentFileCount / FileCount
            ProgressForm.lblCaption.Caption = Str$(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & CurrentFileCount & " of " & FileCount
            ProgressForm.lblFile.Caption = "Processing " & Folder.Files.Item(i)
            ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
            If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
            
            DoEvents ' Necessary for Progress form to update
                
            ' Process each file
            Set File = Filesystem.GetFile(Folder.Files.Item(i))
            Set FileMenu = New Dictionary
            FileMenu.Add "MenuType", "File"
            FileMenu.Add "Name", File.Name
            FileMenu.Add "Path", File.Path
            
            ' Convert headings to bookmarks
            Set Headings = VirtualTub.VTubProcessFile(File.Path)
            FileMenu.Add "Children", Headings
            
            ' Re-initialize the file to get the new modified timestamp
            Set File = Filesystem.GetFile(Folder.Files.Item(i))
            FileMenu.Add "DateLastModified", Format$(File.DateLastModified)
            
            RootMenu.Add Replace(File.Path, "\", "\\"), FileMenu
        End If
    Next
    
    ' Re-initialize the top-level folder for the final modification date
    Set Folder = Filesystem.GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
    RootMenu.Add "DateLastModified", Format$(Folder.DateLastModified)
    
    ' Convert the dictionary to JSON
    JSON = JSONTools.ConvertToJson(RootMenu)
       
    ' Save file
    OutputFile = FreeFile
    Open VTubPath & "VTub.json" For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    Ribbon.RefreshRibbon
    MsgBox "VTub successfully created!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub
    
Handler:
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    Set ProgressForm = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Public Sub VTubRefresh()
    Dim VTubPath As String
    Dim JSON As String
    Dim RootMenu As Object
    Dim Folder As clsFolder
    Dim SubFolder As clsFolder
    Dim FileCount As Long
    Dim Menu As Object
    Dim Children As Object
    Dim Child As Object
    Dim key As Variant
    Dim subkey As Variant
    Dim File As clsFile
    Dim Path As String
    Dim i As Long
    Dim CurrentFileCount As Long
    Dim ProgressPct As Double
    Dim ProgressForm As frmProgress
    Dim OutputFile As Variant
    
    On Error GoTo Handler
    
    ' Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    
    If VTubPath = "" Or VTubPath = "?" Then
        If MsgBox("You haven't configured a folder for the VTub. Open Settings?", vbYesNo, "Open Settings?") = vbYes Then
            UI.ShowForm "Settings"
        End If
        Exit Sub
    End If
    
    If Right$(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Load and parse JSON from file
    JSON = Filesystem.ReadFile(VTubPath & "VTub.json")
    Set RootMenu = JSONTools.ParseJson(JSON)
    
    ' Check file count is the same
    FileCount = 0
    Set Folder = Filesystem.GetFolder(VTubPath)
    
    For i = 1 To Folder.Subfolders.Count
        Set SubFolder = GetFolder(Folder.Subfolders.Item(i))
        FileCount = FileCount + SubFolder.Files.Count
    Next
    
    FileCount = FileCount + Folder.Files.Count
    
    If (CInt(FileCount) <> CInt(RootMenu("FileCount")) Or CInt(Folder.Subfolders.Count) <> CInt(RootMenu("FolderCount"))) Then
        If MsgBox("The number of files or folders in your VTub appear to have changed and needs to be rebuilt from scratch. Rebuild now?", vbOKCancel) = vbCancel Then Exit Sub
        VirtualTub.VTubCreate
        Exit Sub
    End If
    
    CurrentFileCount = 0
    
    ' Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Refreshing VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of approximately " & FileCount
    ProgressForm.Show
    
    ' Iterate each file/folder entry
    For Each key In RootMenu.Keys
        ' Only process File/Folder top level menus
        If key <> "FileCount" And key <> "FolderCount" And key <> "DateLastModified" Then
            Set Menu = RootMenu(key)
            ' Process files in subfolders
            If (Menu("MenuType") = "Folder" And Menu.Exists("Children")) Then
                Set Children = Menu("Children")
                For Each subkey In Children.Keys
                    Set Child = Children(subkey)
                    If Child("MenuType") = "File" Then
                        Path = Child("Path")
                        Set File = Filesystem.GetFile(Path)
                        
                        ' Trap for cancel button on Progress Form
                        If ProgressForm.Visible = False Then Exit Sub
                        
                        ' Update progress form
                        CurrentFileCount = CurrentFileCount + 1
                        ProgressPct = CurrentFileCount / FileCount
                        ProgressForm.lblCaption.Caption = Str$(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & CurrentFileCount & " of " & FileCount
                        ProgressForm.lblFile.Caption = "Processing " & File.Name
                        ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
                        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
                        
                        DoEvents ' Necessary for Progress form to update

                        ' Update the file if the modified date has changed
                        If Child("DateLastModified") <> Format$(File.DateLastModified) Then
                            Set Child("Children") = VirtualTub.VTubProcessFile(File.Path)
                            
                            ' Re-initialize file to get the new modified timestamp
                            Set File = Filesystem.GetFile(File.Path)
                            Child("DateLastModified") = Format$(File.DateLastModified)
                        End If
                    End If
                    
                    Children.Remove subkey
                    Children.Add Replace(subkey, "\", "\\"), Child
                Next subkey
                
                ' Update the subfolder modified date
                Path = Menu("Path")
                Set Folder = Filesystem.GetFolder(Path)
                Menu("DateLastModified") = Format$(Folder.DateLastModified)
            
            ' Process top-level files
            ElseIf Menu("MenuType") = "File" Then
                Path = Menu("Path")
                Set File = Filesystem.GetFile(Path)
                
                ' Trap for cancel button on Progress Form
                If ProgressForm.Visible = False Then Exit Sub
                
                ' Update progress form
                CurrentFileCount = CurrentFileCount + 1
                ProgressPct = CurrentFileCount / FileCount
                ProgressForm.lblCaption.Caption = Str$(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & CurrentFileCount & " of " & FileCount
                ProgressForm.lblFile.Caption = "Processing " & File.Name
                ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
                If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
                
                DoEvents ' Necessary for Progress form to update
            
                ' Update the file if the modified date has changed
                If Menu("DateLastModified") <> Format$(File.DateLastModified) Then
                    Set Menu("Children") = VirtualTub.VTubProcessFile(File.Path)
                    Set File = Filesystem.GetFile(File.Path)
                    Menu("DateLastModified") = Format$(File.DateLastModified)
                End If
            End If
            
            RootMenu.Remove key
            RootMenu.Add Replace(key, "\", "\\"), Menu
        End If
    Next key
    
    ' Update the top-level modification timestamp
    Set Folder = Filesystem.GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
    RootMenu("DateLastModified") = Format$(Folder.DateLastModified)
  
    ' Save new JSON
    JSON = JSONTools.ConvertToJson(RootMenu)
    OutputFile = FreeFile
    Open VTubPath & "VTub.json" For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    ' Refresh ribbon and notify
    Ribbon.RefreshRibbon
    MsgBox "VTub successfully refreshed!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub
    
Handler:
    If Not ProgressForm Is Nothing Then Unload ProgressForm
    Set ProgressForm = Nothing
    
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Private Function VTubProcessFile(ByRef Path As String) As Dictionary
    Dim pCount As Long
    Dim p As Paragraph
    Dim pp As Paragraph
    Dim ppp As Paragraph
      
    Dim Bookmarks As Dictionary
    Set Bookmarks = New Dictionary
    Dim Level1Menu As Dictionary
    Dim Level1Children As Dictionary
    Dim Level2Menu As Dictionary
    Dim Level2Children As Dictionary
    Dim Level3Menu As Dictionary
    Dim StartHeading As Long
    Dim SubHeadingLevel As Long
    Dim Level1Range As Range
    Dim Level2Range As Range
    Dim Level3Range As Range
        
    On Error GoTo Handler
      
    ' Open the file in the background and activate it
    Documents.Open Filename:=Path, Visible:=False
    Documents.Item(Path).Activate
    
    ' Delete all bookmarks
    VirtualTub.RemoveBookmarks
    
    ' Move to top of document and start with largest heading
    pCount = 0
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    StartHeading = Formatting.LargestHeading
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    ' Add a bookmark for every heading level and save it to the menu
    For Each p In Documents.Item(Path).Paragraphs
        pCount = pCount + 1
        If p.OutlineLevel = StartHeading Then
            Set Level1Range = Paperless.SelectHeadingAndContentRange(p)
            Documents.Item(Path).Bookmarks.Add "Level_1_" & pCount, Level1Range
            Set Level1Menu = New Dictionary
            Set Level1Children = New Dictionary
            Level1Menu.Add "Name", "Level_1_" & pCount
            Level1Menu.Add "Path", Path
            Level1Menu.Add "Label", Strings.HeadingToTitle(p.Range.Text)
            Level1Menu.Add "MenuType", "Heading"
            Level1Menu.Add "Children", Level1Children
            Bookmarks.Add "Level_1_" & pCount, Level1Menu
            
            ' Check for nested headings
            SubHeadingLevel = 3
            For Each pp In Level1Range.Paragraphs
                If pp.OutlineLevel = wdOutlineLevel2 Then SubHeadingLevel = 2
            Next pp
            
            For Each pp In Level1Range.Paragraphs
                pCount = pCount + 1
                If pp.OutlineLevel = SubHeadingLevel And pp.OutlineLevel > StartHeading Then
                    Set Level2Range = Paperless.SelectHeadingAndContentRange(pp)
                    Documents.Item(Path).Bookmarks.Add "Level_2_" & pCount, Level2Range
                    Set Level2Menu = New Dictionary
                    Set Level2Children = New Dictionary
                    Level2Menu.Add "Name", "Level_2_" & pCount
                    Level2Menu.Add "Path", Path
                    Level2Menu.Add "Label", Strings.HeadingToTitle(pp.Range.Text)
                    Level2Menu.Add "MenuType", "Heading"
                    Level2Menu.Add "Children", Level2Children
                    Level1Children.Add "Level_2_" & pCount, Level2Menu
                                        
                    If SubHeadingLevel = 2 Then
                        For Each ppp In Level2Range.Paragraphs
                            pCount = pCount + 1
                            If ppp.OutlineLevel = 3 Then
                                Set Level3Range = Paperless.SelectHeadingAndContentRange(ppp)
                                Documents.Item(Path).Bookmarks.Add "Level_3_" & pCount, Level3Range
                                Set Level3Menu = New Dictionary
                                Level3Menu.Add "Name", "Level_3_" & pCount
                                Level3Menu.Add "Label", Strings.HeadingToTitle(ppp.Range.Text)
                                Level3Menu.Add "Path", Path
                                Level3Menu.Add "MenuType", "Heading"
                                Level2Children.Add "Level_3_" & pCount, Level3Menu
                            End If
                        Next ppp
                    End If
                End If
            Next pp
        End If
    Next p
    
    ' Close file and save changes
    Documents.Item(Path).Close SaveChanges:=wdSaveChanges
    
    ' Return the bookmarks
    Set VTubProcessFile = Bookmarks
    
    Exit Function
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Function

Public Sub RemoveBookmarks()
    Dim b As Bookmark
    For Each b In ActiveDocument.Bookmarks
        b.Delete
    Next b
End Sub

