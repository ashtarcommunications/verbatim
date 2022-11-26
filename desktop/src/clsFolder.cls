VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Name As String
Private m_Path As String
Private m_DateLastModified As Date
Private m_Subfolders As New Collection
Private m_Files As New Collection

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = m_DateLastModified
End Property

Public Property Get Subfolders() As Collection
    Set Subfolders = m_Subfolders
End Property

Public Property Get Files() As Collection
    Set Files = m_Files
End Property

Public Sub Init(Path As String)
    Dim Subfolder
    Dim File
    
    m_Path = Path
    
    #If Mac Then
        Dim PathSegments
        PathSegments = Split(Path, Application.PathSeparator)
        m_Name = PathSegments(UBound(PathSegments))
        
        m_DateLastModified = MacScript("do shell script ""stat -f '%m %N' " & Path & "| cut -d ' ' -f 1""")
               
        Dim GetSubfoldersInFolder
        GetSubfoldersInFolder = AppleScriptTask("Verbatim.scpt", "GetSubfoldersInFolder", FolderPath)
        If Right(GetSubfoldersInFolder, 1) = Chr(10) Or Right(GetSubfoldersInFolder, 1) = Chr(13) Then GetSubfoldersInFolder = Left(GetSubfoldersInFolder, Len(GetSubfoldersInFolder) - 1)
        
        For Each Subfolder In GetSubfoldersInFolder
            m_Subfolders.Add Subfolder
        Next Subfolder
           
        POSIXPath = MacScript("tell text 1 thru -2 of " & Chr(34) & Path & Chr(34) & " to return quoted form of it's POSIX Path")
        GetFilesInFolder = AppleScriptTask("Verbatim.scpt", "GetFilesInFolder", POSIXPath)
        If Right(GetFilesInFolder, 1) = Chr(10) Or Right(GetFilesInFolder, 1) = Chr(13) Then GetFilesInFolder = Left(GetFilesInFolder, Len(GetFilesInFolder) - 1)
        
        For Each File In GetFilesInFolder
            m_Files.Add File
        Next File
        
    #Else
        Dim FSO As Scripting.FileSystemObject
        Set FSO = New Scripting.FileSystemObject
        Dim f As Scripting.Folder
        Set f = FSO.GetFolder(Path)
        m_Name = f.Name
        m_DateLastModified = f.DateLastModified
        
        For Each Subfolder In f.Subfolders
            m_Subfolders.Add Subfolder.Path
        Next Subfolder
        
        For Each File In f.Files
            m_Files.Add File.Path
        Next File
        
        Set FSO = Nothing
    #End If
    
    Exit Sub

Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
    #End If
    MsgBox "Error reading " & Path & " - Error " & Err.Number & ": " & Err.Description
End Sub

