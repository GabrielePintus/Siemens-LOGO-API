Attribute VB_Name = "Authentication"
Option Private Module
Dim m_pwd
Dim m_action
Dim m_errorAction
Dim m_arg0
Dim m_connected
Dim m_connectionState
Dim m_connectionListeners ' private data for Connection Listener Module
Dim m_connectionType

Public Const CONNECTION_0BA8 As Integer = 0
Public Const CONNECTION_V81 As Integer = 1


Public Const HTTP_FLAG As Integer = 0
Public Const HTTPS_FLAG As Integer = 1
Public Sub LogIn(ip, pwd, errorAction, action, arg0)
    ' Ensure current is in NOT_CONNECTED state
    LogOut
    
    SetProperty PROPERTY_ID_CONNECTION, STATE_CONNECTING, 0 ' mark it as connecting
    'http first
    'SetHttpFlag HTTP_FLAG
    DoLogIn ip, pwd, errorAction, action, arg0
End Sub

Private Sub DoLogIn(ip, pwd, errorAction, action, arg0)
    ' record user settings
    m_pwd = pwd
    m_action = action
    m_arg0 = arg0
    m_errorAction = errorAction
    
    SetUrl ip
    
    SetRefId Empty ' invalidate the ref id
    SetTempHint Empty
    m_connected = False
    SynchronizeToFile
    'HttpRequest "", "", "OnConnectionFail", "SendUAMCHAL", ""
     HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "OnConnectionFail", "ChallengeMsgHandler", ""
    'If GetHttpFlag() = HTTP_FLAG Then
     '   HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "HttpFailTryHttps", "ChallengeMsgHandler", ""
    'ElseIf GetHttpFlag() = HTTPS_FLAG Then
     '   HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "OnConnectionFail", "ChallengeMsgHandler", ""
    'End If
End Sub

Public Sub LogOut()
    If GetProperty(PROPERTY_ID_CONNECTION) <> STATE_NOT_CONNECTED Then
        m_connected = False
        
        DebugLog "Log Out"
        
        SetProperty PROPERTY_ID_CONNECTION, STATE_NOT_CONNECTED, 0
        
        StopAllRequest ' Clear all on-going requests
        
        If GetRefId() <> 0 Then
            DebugLog "Send Log Out Request"
            
            HttpRequest "AJAX", "UAMLOGOUT:" + CStr(GetRefId()), "OnLogOutFinish", "OnLogOutFinish", ""
        End If
    End If
End Sub

Public Sub OnLogOutFinish(arg0, status)

    HttpRequest "abort", "", "", "", ""
End Sub

Public Sub RecoverConnection()
    ' Try recovery only in Connected State
    If GetProperty(PROPERTY_ID_CONNECTION) = STATE_CONNECTED Then
        DebugLog "Recovery"
        
        SetProperty PROPERTY_ID_CONNECTION, STATE_RECOVERING, 0
        
        StopAllRequest ' Clear all on-going requests
        
        DoLogIn GetUrl(), m_pwd, "CommonOnConnectionRecoverFail", "", ""
    End If
End Sub

Public Function GetConnectionType()
    GetConnectionType = m_connectionType
End Function

Public Sub OnConnectionFail(arg0, status)
    'MsgBox "Connection Error:" + CStr(status)
    
    ReportLoginError status
End Sub


Public Sub HttpFailTryHttps(arg0, status)
    'MsgBox "Connection Error:" + CStr(status)
    m_connected = False
    LogOut
    SetHttpFlag (HTTPS_FLAG)
    HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "OnConnectionFail", "ChallengeMsgHandler", ""
End Sub

Private Sub SendUAMCHAL(arg0, data As String)
    SynchronizeToFile
    DebugLog "[Authentication] Login connect is ok, begin SendUAMCHAL "
    HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "OnConnectionFail", "ChallengeMsgHandler", ""
End Sub

Private Sub ChallengeMsgHandler(arg0, data As String)
    Dim parts
    parts = Split(data, ",")
    
    If UBound(parts) >= 1 Then
        If CInt(parts(0)) = 700 Then
            Dim ref
            Dim key2
            
            If UBound(parts) = 2 Then
                ref = parts(1)
                key2 = CDbl(parts(2))
            End If
        
            Dim serverChanllengeToke
            serverChanllengeToke = CalculateXOR(CalculateXOR(CalculateXOR(CalculateXOR(10, 10), 10), 10), key2)
            
            Dim passwordToken
            passwordToken = CalculateXOR(MakeCRC32(String2UTF8(m_pwd + "+" + parts(2))), key2)
            
            SetTempHint ref
            
            HttpRequest "AJAX", "UAMLOGIN:Web User," + CStr(passwordToken) + "," + CStr(serverChanllengeToke), _
                "OnConnectionFail", "LogInMsgHandler", ""
        Else
            ' login error happens
        End If
    Else
        ' unexpected error
    End If
End Sub

Private Sub LogInMsgHandler(arg0, data As String)
    Dim parts
    parts = Split(data, ",")
    
    Dim loginStatus
    loginStatus = 603 ' logical error by default
    
    If UBound(parts) >= 1 Then
        If CInt(parts(0)) = 700 Then
            If UBound(parts) = 1 Then
                loginStatus = 700
                SetRefId parts(1)
            End If
        Else
            ' login error happens
            loginStatus = CInt(parts(0))
        End If
    Else
        ' unexpected error
    End If
    
    If loginStatus = 700 Then
        HttpRequest "AJAX", "GETFWVER", "OnConnectionFail", "SystemInfoHandler", ""
    Else
        ReportLoginError loginStatus
    End If
End Sub

' The final step of Log In -- Get System Info
Private Sub SystemInfoHandler(arg0, data As String)
    If data = "" Then
        m_connectionType = CONNECTION_0BA8
        'm_connectionType = CONNECTION_V81 ' It is V8.1 anyway
    Else
        m_connectionType = CONNECTION_V81
    End If
    
    ' notify user application about login result
    m_connected = True
    
    SetProperty PROPERTY_ID_CONNECTION, STATE_CONNECTED, 0
    If m_action <> "" Then
        Application.Run m_action, m_arg0, 700
    End If
End Sub

Private Sub ReportLoginError(code)
    m_connected = False
    
    ' Don't report error when it is already in STATE_NOT_CONNECTED state
    If GetProperty(PROPERTY_ID_CONNECTION) <> STATE_NOT_CONNECTED Then
        ' let handler handle the case before change property then it could
        ' know its current status
        If m_errorAction <> "" Then
            Application.Run m_errorAction, m_arg0, code
        End If
        
        SetProperty PROPERTY_ID_CONNECTION, STATE_NOT_CONNECTED, code
    End If
End Sub
