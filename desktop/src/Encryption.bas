Attribute VB_Name = "Encryption"
Option Explicit

Public Function XORDecryption(DataIn As String) As String
    
    Dim CodeKey As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer
    
    'Generate a unique CodeKey
    CodeKey = GetHDSerial()
    
    For lonDataPtr = 1 To (Len(DataIn) / 2)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Val("&H" & (Mid$(DataIn, (2 * lonDataPtr) - 1, 2)))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr(intXOrValue1 Xor intXOrValue2)
    Next lonDataPtr
    XORDecryption = strDataOut
End Function

Public Function XOREncryption(DataIn As String) As String
    
    Dim CodeKey As String
    
    Dim lonDataPtr As Long
    Dim strDataOut As String
    Dim temp As Integer
    Dim tempstring As String
    Dim intXOrValue1 As Integer
    Dim intXOrValue2 As Integer

    'Generate a unique CodeKey
    CodeKey = GetHDSerial()
    
    For lonDataPtr = 1 To Len(DataIn)
        'The first value to be XOr-ed comes from the data to be encrypted
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        'The second value comes from the code key
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        temp = (intXOrValue1 Xor intXOrValue2)
        tempstring = Hex(temp)
        If Len(tempstring) = 1 Then tempstring = "0" & tempstring
        
        strDataOut = strDataOut + tempstring
    Next lonDataPtr
    XOREncryption = strDataOut
End Function

Private Function GetHDSerial() As String
'Generates a unique computer ID using the c:\ volume serial.
'Fails if there's no c drive for some reason

    Dim Serial As Variant
    Dim FSO As FileSystemObject
    Set FSO = New FileSystemObject
    
    'Turn off error checking - if anything goes wrong it will use a default value
    On Error Resume Next
    
    'Get the Volume Serial for the c drive
    Serial = FSO.GetDrive("c").SerialNumber
   
    'Convert to hex, triple for length
    Serial = Hex(Serial)
    Serial = Serial & Serial & Serial
    
    'If something went wrong above or a real number wasn't returned, set a default
    If Len(Serial) < 3 Then
        Serial = "dj2ijg84nvnwj38gnm90dopqm9256dmn"
    End If
    
    'Close FSO
    Set FSO = Nothing
    
    'Set return value
    GetHDSerial = Serial

End Function


