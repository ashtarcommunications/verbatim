Attribute VB_Name = "HTTP"
Option Explicit

Public Function GetReq(URL As String) As Dictionary
    On Error GoTo Handler
    
    Dim Response As Dictionary
    Set Response = New Dictionary
        
    #If Mac Then
        Dim Script As String
        Dim Raw As String
        Dim StatusCode As String
        Dim Body As String
        
        ' Get the response from curl as <status_code>\n<body> by redirecting stderr to stdout
        ' -w is the write-out string format, -o redirects to standard, -s is silent
        Script = "curl -w '%{stderr}%{http_code}\n%{stdout}' -s -o - '" & URL & "' 2>&1"
        Raw = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
        
        ' Chr(13) = matches \n newline output from curl - should only be one newline in output
        StatusCode = Split(Raw, Chr(13))(0)
        Body = Split(Raw, Chr(13))(1)

        Response.Add "status", StatusCode
        Response.Add "body", JSONTools.ParseJson(Body)
        
        Set GetReq = Response
        
        Set Response = Nothing
    #Else
        Dim HttpReq As MSXML2.ServerXMLHTTP60
        Set HttpReq = New ServerXMLHTTP60
        HttpReq.setTimeouts 10000, 10000, 30000, 30000
        HttpReq.Open "GET", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send
        
        Response.Add "status", HttpReq.Status
        Response.Add "body", JSONTools.ParseJson(HttpReq.responseText)
             
        Set GetReq = Response

        Set HttpReq = Nothing
        Set Response = Nothing
    #End If

    Exit Function
    
Handler:
    ' Return an empty response if there was a network error
    Response.Add "status", "Error " & Err.Number & ": " & Err.Description
    Response.Add "body", JSONTools.ParseJson("[]")
    Set GetReq = Response
    
    #If Mac Then
        Set Response = Nothing
    #Else
        Set HttpReq = Nothing
        Set Response = Nothing
    #End If
End Function

Public Function PostReq(URL As String, Body As Dictionary) As Dictionary
    On Error GoTo Handler
    
    Dim Response As Dictionary
    Set Response = New Dictionary
        
    Dim JSON
    JSON = JSONTools.ConvertToJson(Body)
    
    #If Mac Then
        Dim Script As String
        Dim Cookie As String
        Dim Raw As String
        Dim StatusCode As String
        Dim ResponseBody As String
        
        Cookie = GetSetting("Verbatim", "Caselist", "CaselistToken", vbNullString)
        
        ' Uses same output redirection as GET to retrieve status code and response body
        Script = "curl -X POST "
        Script = Script & "-H 'Content-Type: application/json' "
        Script = Script & "-H 'Cookie: caselist_token=" & Cookie & "' "
        Script = Script & "-d '" & JSON & "' "
        Script = Script & "-w '%{stderr}%{http_code}\n%{stdout}' -s -o - "
        Script = Script & "'" & URL & "'"
        Script = Script & " 2>&1"
        
        Raw = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
        StatusCode = Split(Raw, Chr(13))(0)
        ResponseBody = Split(Raw, Chr(13))(1)

        Response.Add "status", StatusCode
        Response.Add "body", JSONTools.ParseJson(ResponseBody)
        
        Set PostReq = Response
        
        Set Response = Nothing
    #Else
        Dim HttpReq As MSXML2.ServerXMLHTTP60
        Set HttpReq = New ServerXMLHTTP60
        HttpReq.setTimeouts 10000, 10000, 30000, 30000
        HttpReq.Open "POST", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Content-Type", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send JSON
        
        Response.Add "status", HttpReq.Status
        Dim ResponseBody As Object
        Set ResponseBody = JSONTools.ParseJson(HttpReq.responseText)
        
        Response.Add "body", ResponseBody
      
        Set PostReq = Response
      
        Set HttpReq = Nothing
        Set Response = Nothing
        Set ResponseBody = Nothing
    #End If

    Exit Function
    
Handler:
    ' Return an empty response if there was a network error
    Response.Add "status", "Error " & Err.Number & ": " & Err.Description
    Response.Add "body", JSONTools.ParseJson("[]")
    Set PostReq = Response
    
    #If Mac Then
        Set Response = Nothing
    #Else
        Set HttpReq = Nothing
        Set Response = Nothing
        Set ResponseBody = Nothing
    #End If
End Function


'Public Sub DownloadFile(URL As String, FilePath As Variant)
'    #If Mac Then
'        Dim Script As String
'        Script = "curl - o '" & URL & "'" & " " & FilePath & """"
'        AppleScriptTask "Verbatim.scpt", "RunShellScript", Script
'    #Else
'        Dim HttpReq As MSXML2.ServerXMLHTTP60
'        Set HttpReq = New ServerXMLHTTP60
'        HttpReq.Open "GET", URL, False
'        HttpReq.send
'        Dim FileStream
'        Set FileStream = CreateObject("ADODB.Stream")
'        FileStream.Open
'        FileStream.Type = 1
'        FileStream.Write HttpReq.ResponseBody
'        FileStream.SaveToFile FilePath, 2 '1 = no overwrite, 2 = overwrite
'        FileStream.Close
'        Set FileStream = Nothing
'    #End If
'End Sub
