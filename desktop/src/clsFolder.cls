VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Folder class for a cross-platform abstraction to mimic the FileSystemObject on Mac
' Does not implement the complete Folder API
' Uses Scripting.FileSystemObject on PC, shell scripting (via AppleScriptTask) on Mac
'
' Usage:
' Dim f as clsFolder
' Set f = New clsFolder
' f.Init <PathToFolder>

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

Public Sub Init(ByRef Path As String)
    Dim SubFolder As Variant
    Dim File As Variant
    
    On Error GoTo Handler
    
    m_Path = Path
    
    #If Mac Then
        Dim PathSegments
        Dim Script As String
        Dim Subfolders
        Dim Files
        
        PathSegments = Split(Path, Application.PathSeparator)
        m_Name = PathSegments(UBound(PathSegments))
                
        Script = "stat -f '%m' '" & Path & "'"
        m_DateLastModified = Strings.ConvertUnixTimestampToDate(AppleScriptTask("Verbatim.scpt", "RunShellScript", Script))

        Script = "find '" & Path & "' -type d -depth 1"
        Subfolders = Split(AppleScriptTask("Verbatim.scpt", "RunShellScript", Script), Chr(13))
        
        For Each SubFolder In Subfolders
            m_Subfolders.Add Replace(SubFolder, "//", "/")
        Next SubFolder
        
        Script = "find '" & Path & "' -type f -depth 1"
        Files = Split(AppleScriptTask("Verbatim.scpt", "RunShellScript", Script), Chr(13))
        
        For Each File In Files
            m_Files.Add Replace(File, "//", "/")
        Next File
        
    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Dim f As Object
        Set f = FSO.GetFolder(Path)
        m_Name = f.Name
        m_DateLastModified = f.DateLastModified
        
        For Each SubFolder In f.Subfolders
            m_Subfolders.Add SubFolder.Path
        Next SubFolder
        
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

