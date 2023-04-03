VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
' Menu class for building a nested dynamic VTub Menu in JSON
' MenuType is either a File or a Folder, each menu will have nested children
'
' The ToJSON method serializes the class for storage in VTub.json

Option Explicit

Private m_MenuType As String
Private m_Name As String
Private m_Path As String
Private m_DateLastModified As Date

Public Property Get MenuType() As String
    MenuType = m_MenuType
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Get DateLastModified() As Date
    DateLastModified = m_DateLastModified
End Property

Public Sub Init(ByVal MenuType As String, ByVal Name As String, ByVal Path As String, ByVal DateLastModified As Date)
    On Error GoTo Handler

    m_MenuType = MenuType
    m_Name = Name
    m_Path = Path
    m_DateLastModified = DateLastModified
    
    Exit Sub

Handler:
    MsgBox "Error reading " & Path & " - Error " & Err.Number & ": " & Err.Description
End Sub

Public Function ToJSON() As String
    ToJSON = "{""MenuType"":"""
    ToJSON = ToJSON & m_MenuType
    ToJSON = ToJSON & """, ""Name"":"""
    ToJSON = ToJSON & m_Name
    ToJSON = ToJSON & """, ""Path"":"""
    ToJSON = ToJSON & m_Path
    ToJSON = ToJSON & """, ""DateLastModified"":"""
    ToJSON = ToJSON & m_DateLastModified
    ToJSON = ToJSON & """}"
End Function

