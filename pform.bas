Attribute VB_Name = "Module1"
'Option Explicit
DefLng A-Z

Const INTERNET_INVALID_PORT_NUMBER = 0
Const INTERNET_SERVICE_HTTP = 3
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_KEEP_CONNECTION = &H400000

Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sbuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Declare Function HttpSendRequest Lib _
"wininet.dll" Alias "HttpSendRequestA" _
(ByVal hHttpRequest As Long, ByVal sHeaders _
As String, ByVal lHeadersLength As Long, _
ByVal sOptional As String, ByVal _
lOptionalLength As Long) As Integer

Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

Global m_cPostBuffer$


'Dim m_cPostBuffer As String

Public Function PostForm(ByVal server As String, ByVal CGI As String) As String
   
    On Error GoTo myError
    Dim hOpen As Long, hConnection As Long
    Dim hURL As Long
    Dim sbuffer As String
    Dim lNumBytesToRead  As Long
    Dim lNumberOfBytesRead As Long
    Dim Result As String
    
    ' open internet connection
    hOpen = InternetOpen(App.EXEName, INTERNET_OPEN_TYPE_PRECONFIG, "", "", 0)
    If hOpen <> 0 Then
        hConnection = InternetConnect(hOpen, server, INTERNET_INVALID_PORT_NUMBER, "", "", INTERNET_SERVICE_HTTP, 0, 0)
        If hConnection <> 0 Then
            hURL = HttpOpenRequest(hConnection, "POST", CGI, "", "", 0, INTERNET_FLAG_KEEP_CONNECTION, 0)
            If hURL <> 0 Then
                If HttpSendRequest(hURL, "", 0, m_cPostBuffer, Len(m_cPostBuffer)) Then
                    lNumBytesToRead = 1024
                    sbuffer = Space$(lNumBytesToRead)
                    Do While InternetReadFile(hURL, sbuffer, lNumBytesToRead, lNumberOfBytesRead)
                        If lNumberOfBytesRead = 0 Then
                            Exit Do
                        Else
                            Result = Result & Left$(sbuffer, lNumberOfBytesRead)
                        End If
                        lNumBytesToRead = 1024
                        sbuffer = Space$(lNumBytesToRead)
                    Loop
                    PostForm = Trim$(Result)
                Else
                    Err.Raise vbObjectError + 504, , "HttpSendRequest"
                End If
            Else
                Err.Raise vbObjectError + 505, , "HttpOpenRequest"
            End If
        Else
            Err.Raise vbObjectError + 506, , "InternetConnect"
        End If
    Else
        Err.Raise vbObjectError + 507, , "InternetOpen"
    End If
    
myExit:
    InternetCloseHandle hURL
    InternetCloseHandle hConnection
    InternetCloseHandle hOpen
    Exit Function

myError:
    PostForm = "ERROR " & Err.Description
    Resume myExit
    
End Function

Private Function UrlEncode(sText As String) As String
    
    Dim sResult As String
    Dim sFinal As String
    Dim sChar As String
    Dim i As Long
    
    For i = 1 To Len(sText)
        
        sChar = Mid$(sText, i, 1)
        
        If InStr(1, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789", sChar) <> 0 Then
            sResult = sResult & sChar
        ElseIf sChar = " " Then
            sResult = sResult & "+"
        ElseIf True Then
            sResult = sResult & "%" & Right$("0" & Hex(Asc(sChar)), 2)
        End If
        
        If Len(sResult) > 1000 Then
            sFinal = sFinal & sResult
            sResult = ""
        End If
    
    Next
    
    UrlEncode = sFinal & sResult

End Function


Public Function AddPostKey(tckey As String, tcValue As String)
    m_cPostBuffer = m_cPostBuffer & UrlEncode(tckey) & _
                       "=" + UrlEncode(tcValue) '+ "&"
End Function


