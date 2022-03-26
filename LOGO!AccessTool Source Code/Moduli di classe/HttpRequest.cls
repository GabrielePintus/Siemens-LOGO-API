VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HttpRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_httpReq As Object
Dim m_errorAction
Dim m_msgAction
Dim m_arg0
Dim m_status As Integer ' 0: normal, 1: requesting, 2: requested
Dim m_locked As Boolean
Dim m_executionTime As Integer
Dim m_id
Dim m_retry
Dim m_url
Dim m_data

Const MAX_EXECUTION_TIME As Integer = 6

Const REQUEST_STATE_NORMAL As Integer = 0
Const REQUEST_STATE_REQUESTING As Integer = 1
Const REQUEST_STATE_FINISHED As Integer = 2

Private Sub Class_Initialize()
    NewHttpReq
    
    m_status = REQUEST_STATE_NORMAL
End Sub

Private Sub NewHttpReq()
    If Not m_httpReq Is Nothing Then
        m_httpReq.onreadystatechange = Nothing
        Set m_httpReq = Nothing
    End If
    
    'Set m_httpReq = CreateObject("Msxml2.ServerXMLHTTP.6.0")
     Set m_httpReq = CreateObject("MSXML2.XMLHTTP.6.0")
     
    'm_httpReq.setOption 2, 13056 'SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
     'm_httpReq.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS

    'm_httpReq.Open "get", "https://" + m_ip + "/" + url, True
    'm_httpReq.Open "get", "https://192.168.0.53/", True
    'm_httpReq.Send
    m_httpReq.onreadystatechange = Me
End Sub

Public Sub SetId(id)
    m_id = id
End Sub

Public Function GetId()
    GetId = m_id
End Function

Sub Initialize(url, data, errorAction, msgAction, arg0, retry)
    m_url = url
    m_data = data
    m_errorAction = errorAction
    m_msgAction = msgAction
    m_arg0 = arg0
    m_retry = retry
    
    DebugLog "[HttpRequest] (" + CStr(m_id) + ") Initialized ReadyState " + CStr(m_httpReq.readyState)
    
    m_status = REQUEST_STATE_REQUESTING
    m_executionTime = 0 ' Reset execution time when initialized
    
    SendHttpData m_httpReq, m_url, m_data
End Sub

Public Function GetHttp()
    Set GetHttp = m_httpReq
End Function

Sub ReadyStateChangeHandler()
Attribute ReadyStateChangeHandler.VB_UserMemId = 0
    On Error GoTo Err
    ' DebugLog "[HttpRequest] ReadyStateChange " + CStr(m_id) + " Locked " + CStr(m_locked) + " ReadyState " + CStr(m_httpReq.readyState)
    
    If m_locked Then
        Exit Sub
    End If
    
    m_locked = True
    
    ' directly pass httpReq as arg is acceptable, do it if necessary
    'MsgBox m_httpReq.readyState
    
    If m_httpReq.readyState = 4 Then
        m_status = REQUEST_STATE_FINISHED ' State: Requested
        
        HandleFinishedRequest
        'TaskHandle ' This requires additional time for handling
        
        'm_httpReq.onreadystatechange = Empty
        'm_httpReq.onreadystatechange = Nothing
        'm_httpReq = Empty
     
      Else
    
    End If
    
    ' General operation for normal operation / error handling
Err:
    m_locked = False
    
    If m_status = REQUEST_STATE_NORMAL Then
        StartReq
    End If
End Sub

Public Sub TickExecutionTime()
    If m_status = REQUEST_STATE_REQUESTING Then
        m_executionTime = m_executionTime + 1
    End If
End Sub

Private Sub AbortHttpRequest()
    m_httpReq.abort
    NewHttpReq
End Sub

Public Function IsAvailable()
    On Error GoTo Err

    If m_locked Then
        IsAvailable = False
        Exit Function
    End If
    
    m_locked = True
    
    If m_status = REQUEST_STATE_REQUESTING Then
        ' try to expire requesting
        If m_executionTime > MAX_EXECUTION_TIME Then
            DebugLog "[HttpRequest] (" + CStr(m_id) + ") Expired ReadyState " + CStr(m_httpReq.readyState)
        
            AbortHttpRequest
            
            ' For a timeout request. It got the chance to retry
            If m_retry > 0 Then
                m_retry = m_retry - 1
                DebugLog "[HttpRequest] (" + CStr(m_id) + ") Retry"
        
                m_status = REQUEST_STATE_REQUESTING
                m_executionTime = 0 ' Reset execution time when initialized
                
                SendHttpData m_httpReq, m_url, m_data
            Else
                RunApp m_errorAction, m_arg0, 101 ' TODO: use http-timeout as error code
            
                ResetRequest
            End If
        End If
    Else
        If m_status = REQUEST_STATE_FINISHED Then
            ' try to handle requested
            HandleFinishedRequest
        End If
    End If

Err:
    m_locked = False
    
    If m_status = REQUEST_STATE_NORMAL Then
        IsAvailable = True
    Else
        IsAvailable = False
    End If
End Function

Public Sub Cancel()
    On Error GoTo Err
    
    If m_locked Then
        Exit Sub
    End If
    
    m_locked = True
    
    If m_status = REQUEST_STATE_REQUESTING Then
        DebugLog "[HttpRequest] (" + CStr(m_id) + ") Cancel ReadyState " + CStr(m_httpReq.readyState)
    
        AbortHttpRequest
        
        'DebugLog "[HttpRequest] (" + CStr(m_id) + ") Cancelled ReadyState " + CStr(m_httpReq.readyState)
        
        RunApp m_errorAction, m_arg0, 0 ' TODO: use http-timeout as error code
        
        ResetRequest
    End If
    
Err:
    m_locked = False
    
    If m_status = REQUEST_STATE_NORMAL Then
        StartReq
    End If
End Sub

Private Sub HandleFinishedRequest()
    DebugLog "[HttpRequest] (" + CStr(m_id) + ") Complete " + CStr(m_httpReq.status)
        
    If m_httpReq.status = 200 Then
        RunApp m_msgAction, m_arg0, m_httpReq.responseText
    'ElseIf m_httpReq.status <> 0 Then
     '   MsgBox m_httpReq.status
    ElseIf GetHttpFlag() = HTTP_FLAG Then
        SetHttpFlag (HTTPS_FLAG)
    HttpRequest "AJAX", "UAMCHAL:3,4,10,10,10,10", "OnConnectionFail", "ChallengeMsgHandler", ""
    
    Else
        RunApp m_errorAction, m_arg0, m_httpReq.status
    End If
    
    ResetRequest
End Sub

Private Sub ResetRequest()
    m_status = REQUEST_STATE_NORMAL
    m_executionTime = 0
End Sub

Public Sub RunApp(action, arg0, arg1)
    If action <> "" Then
        Application.Run action, arg0, arg1
    End If
End Sub

Private Sub TaskHandle()
    DebugLog "[HttpRequest] TaskHandle"
    Application.OnTime Now + TimeValue("00:00:00"), "StartReq" ' Try to start handling in timer thread
End Sub

