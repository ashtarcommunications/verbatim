VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' File class for a cross-platform abstraction to mimic the FileSystemObject on Mac
' Does not implement the complete File API
' Uses Scripting.FileSystemObject on PC, shell scripting (via AppleScriptTask) on Mac
'
' Usage:
' Dim f as clsFile
' Set f = New clsFile
' f.Init <PathToFile>

Option Explicit

Private m_Name As String
Private m_Path As String
Private m_DateLastModified As Date

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = m_DateLastModified
End Property

Public Sub Init(ByRef Path As String)
    On Error GoTo Handler
    
    m_Path = Path
    
    #If Mac Then
        Dim PathSegments
        PathSegments = Split(Path, Application.PathSeparator)
        m_Name = PathSegments(UBound(PathSegments))
        
        Dim Script As String
        Script = "stat -f '%m' '" & Path & "'"
        m_DateLastModified = Strings.ConvertUnixTimestampToDate(AppleScriptTask("Verbatim.scpt", "RunShellScript", Script))

    #Else
        Dim FSO As Object
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Dim f As Object
        Set f = FSO.GetFile(Path)
        m_Name = f.Name
        m_DateLastModified = f.DateLastModified
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
