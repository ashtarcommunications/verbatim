Attribute VB_Name = "Audio"
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage, VariableNotAssigned
' Adapated from http://www.rediware.com/programming/vb/vbrecwav/vbrecordwav.htm
Option Explicit

Public Sub StartRecord(ByRef Channels As String, ByRef Bits As String, ByRef Samples As String)
    Dim ReturnString As String * 1024
    
    On Error GoTo Handler
        
    ' Save instead if already recording
    If Audio.RecordStatus = "recording" Then Audio.SaveRecord
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "StartRecord", ""
    #Else
        ' Start recording
        mciSendString "open new Type waveaudio Alias recsound", ReturnString, Len(ReturnString), 0
        mciSendString "set recsound time format ms", ReturnString, 1024, 0
        mciSendString "set recsound format tag pcm", ReturnString, 1024, 0
        mciSendString "set recsound channels " & Channels, ReturnString, 1024, 0
        mciSendString "set recsound samplespersec " & Samples, ReturnString, 1024, 0
        mciSendString "set recsound bitspersample " & Bits, ReturnString, 1024, 0
        mciSendString "set recsound alignment " & Str$(CInt((CLng(Bits) / 8) * CLng(Channels))), ReturnString, 1024, 0
        mciSendString "record recsound", ReturnString, Len(ReturnString), 0
        
        ' Save bitrate for fixing wav file
        Globals.BytesPerSec = CLng(Samples) * ((CLng(Channels) * CLng(Bits)) / 8)
    #End If
    
    MsgBox "Recording started. Press the button again to stop."
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Audio.RecordStatus = "recording" Then
        Audio.SaveRecord
    End If
    Globals.RecordAudioToggle = False
    Ribbon.RefreshRibbon
    
End Sub

Public Sub SaveRecord()
    Dim ReturnString As String * 1024
    
    Dim AudioDir As String
    Dim Filename As String
    Dim FilePath As String
    
    On Error GoTo Handler
    
    ' Get Audio recording directory from settings or use desktop
    AudioDir = GetSetting("Verbatim", "Paperless", "AudioDir", "")
    If Filesystem.FolderExists(AudioDir) = False Then
        #If Mac Then
            FilePath = "/Users/" & Environ("USER") & "/Desktop"
        #Else
            FilePath = CStr(Environ$("USERPROFILE")) & "\Desktop"
        #End If
    Else
        FilePath = AudioDir
    End If

    ' Ensure a trailing separator
    If Right$(FilePath, 1) <> Application.PathSeparator Then FilePath = FilePath & Application.PathSeparator
             
GetFileName:
    Filename = InputBox("Please enter a name for your saved audio file. It will be saved to the following directory:" _
        & vbCrLf & "(Configurable in Settings)" & vbCrLf & FilePath, _
        "Save Audio Recording", _
        "Recording " & Format$(Now, "m d yyyy hmmAMPM"))
    
    ' Exit if no file name or user pressed Cancel, recording is still active
    If Filename = "" Then
        Globals.RecordAudioToggle = True
        Exit Sub
    End If
    
    ' Clean up filename and ensure correct extension
    Filename = Strings.OnlyAlphaNumericChars(Filename)
    #If Mac Then
        If Right(Filename, 4) <> ".m4a" Then Filename = Filename & ".m4a"
    #Else
        If Right$(Filename, 4) <> ".wav" Then Filename = Filename & ".wav"
    #End If
    Filename = FilePath & Filename
    
    ' Test if file exists
    If Filesystem.FileExists(Filename) = True Then
        If MsgBox("File exists. Overwrite?", vbYesNo) = vbNo Then GoTo GetFileName
    End If
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "SaveRecord", Filename
    #Else
        ' Stop and save capture
        mciSendString "stop recsound", ReturnString, Len(ReturnString), 0
        mciSendString "save recsound " & """" & Filename & """", ReturnString, Len(ReturnString), 0
        mciSendString "close recsound", ReturnString, 1024, 0
        mciSendString "close all", 0, 0, 0
        
        ' Fix mciSendString bug saving the wrong bitrate
        Audio.FixWavFile Filename
    #End If
    
    MsgBox "Recording Saved as:" & vbCrLf & Filename, vbOKOnly
    Ribbon.RefreshRibbon
    
    Exit Sub
    
Handler:
    RecordAudioToggle = False
    #If Mac Then
        ' Do Nothing
    #Else
        If Audio.RecordStatus = "recording" Then
            mciSendString "stop recsound", ReturnString, Len(ReturnString), 0
        End If
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Ribbon.RefreshRibbon
End Sub

Public Function RecordStatus() As String
    #If Mac Then
        If Globals.RecordAudioToggle = True Then
            RecordStatus = "recording"
        Else
            RecordStatus = "off"
        End If
    #Else
        Dim ReturnString As String * 1024
        mciSendString "status recsound mode", ReturnString, Len(ReturnString), 0
        RecordStatus = ReturnString
    #End If
End Function

'@Ignore ParameterNotUsed
Public Sub RecordAudio(ByVal c As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo Handler
    
    If pressed Then
        ' Start recording - Mac ignores the parameters
        Audio.StartRecord "1", "8", "44100"
        Globals.RecordAudioToggle = True
    Else
        Globals.RecordAudioToggle = False
        
        ' Stop and save recording
        Audio.SaveRecord
    End If
    
    Ribbon.RefreshRibbon
    
    Exit Sub
    
Handler:
    Globals.RecordAudioToggle = False
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Ribbon.RefreshRibbon
End Sub

Public Sub FixWavFile(ByVal FilePath As String)
    '@Ignore IntegerDataType
    Dim Indexnum As Integer
    Dim HexCode As String
    Dim Hex1 As String
    Dim Hex2 As String
    Dim Hex3 As String
    Dim lByteNum As Long
    Dim bByte As Byte
    
    HexCode = Hex$(Globals.BytesPerSec)
    Do While Len(HexCode) < 6
        HexCode = "0" & HexCode
    Loop
    
    Hex1 = Right$(HexCode, 2)
    Hex2 = Mid$(HexCode, 3, 2)
    Hex3 = Left$(HexCode, 2)
    
    ' Manually fix the inaccurate bitrate in hex
    Indexnum = FreeFile
    Open FilePath For Binary Access Write As #Indexnum
    lByteNum = 29
    bByte = CInt("&H" & Hex1)
    Put #Indexnum, lByteNum, bByte
    bByte = CInt("&H" & Hex2)
    lByteNum = lByteNum + 1
    Put #Indexnum, lByteNum, bByte
    bByte = CInt("&H" & Hex3)
    lByteNum = lByteNum + 1
    Put #Indexnum, lByteNum, bByte
    Close #Indexnum
End Sub


