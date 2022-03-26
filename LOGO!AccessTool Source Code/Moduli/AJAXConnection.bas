Attribute VB_Name = "AJAXConnection"
Option Private Module
Dim m_firstReq As AJAXReq
Dim m_lastReq As AJAXReq
Dim m_ip As String
Dim m_ref As String  'in the LOGO8.1 service pack version, BM changed ,the m_ref should be string instead of integer
Dim m_tempHint As String
Dim m_operatingReq(0) As HttpRequest ' assuming 10 is enough. (It can be changed by InitializeOperatingReq)
Dim m_refreshing As Boolean
Dim m_count As Integer
Dim m_httpflag As Integer


Public Sub EnqueueReq(req As AJAXReq)
    If m_firstReq Is Nothing Then
        Set m_firstReq = req
        Set m_lastReq = req
    Else
        m_lastReq.SetNext req
        Set m_lastReq = req
    End If
    
    m_count = m_count + 1
End Sub

Public Sub TestQueue()
    MsgBox m_count
End Sub

Public Function IsQueueEmpty()
    If m_firstReq Is Nothing Then
        IsQueueEmpty = True
    Else
        IsQueueEmpty = False
    End If
End Function

Public Function DequeueReq()
    If m_firstReq Is Nothing Then
        Set DequeueReq = Nothing
    Else
        Dim dequedReq
        
        Set dequedReq = m_firstReq
        
        If m_firstReq Is m_lastReq Then
            Set m_firstReq = Nothing
            Set m_lastReq = Nothing
        Else
            Set m_firstReq = m_firstReq.GetNext()
        End If
        
        Set DequeueReq = dequedReq
        
        m_count = m_count - 1
    End If
End Function

Public Sub SetUrl(ip)
    m_ip = ip
End Sub

Public Function GetUrl()
    GetUrl = m_ip
End Function

Public Sub SetHttpFlag(httpflag)
    m_httpflag = httpflag
End Sub

Public Function GetHttpFlag()
    GetHttpFlag = m_httpflag
End Function

Public Sub HttpRequest(url, data, errorAction, msgAction, arg0)
    HttpTimeoutRequest url, data, errorAction, msgAction, arg0, 0 ' No retry by default
End Sub

Public Sub HttpTimeoutRequest(url, data, errorAction, msgAction, arg0, retry)
    Dim req As AJAXReq
    
    Set req = New AJAXReq
    req.Initialize url, data, errorAction, msgAction, arg0
    req.SetRetry retry
    
    EnqueueReq req
    
    StartReq
    
    TryStartTimeout
End Sub

Public Sub StopAllRequest()
    DebugLog "[AJAXConnection] Stop All"
    
    ' Cancel all pending request
    Do While Not IsQueueEmpty()
        Set req = DequeueReq()
            
        ' Double Check req Because it might be called simultaneously.
        ' Remarks: It could also be Nothing because GetAvailableHttpRequest
        ' might have been triggered reschedule already.
        If Not req Is Nothing Then
            req.Cancel
        End If
    Loop
    
    ' Cancel all on-going request
    Dim i
    For i = LBound(m_operatingReq) To UBound(m_operatingReq) Step 1
        If Not m_operatingReq(i) Is Nothing Then
            'DebugLog "[AJAXConnection] " + CStr(i) + " ReadyState " + CStr(m_operatingReq(i).GetHttp.readyState)

            m_operatingReq(i).Cancel
        End If
    Next
End Sub

Public Sub SetRefId(ref)
    m_ref = ref
    DebugLog "[AJAXConnection] NewRefId " + ref
End Sub

Public Function GetRefId()
    GetRefId = m_ref
End Function

Public Sub SetTempHint(hint)
    m_tempHint = hint
End Sub

Public Sub SendHttpData(httpReq, url, data)
    ' Make a request only when ip is valid. Otherwise do nothing to make it timeout
    'the tcp could not be closed by httpreq now.
    If url = "abort" Then
        httpReq.abort
        Exit Sub
    End If
    If m_ip <> "" Then
        If m_httpflag = HTTP_FLAG Then
            DebugLog "[AJAXConnection] Encrymode:http"
            httpReq.Open "post", "http://" + m_ip + "/" + url, True
        ElseIf m_httpflag = HTTPS_FLAG Then
            DebugLog "[AJAXConnection] Encrymode: secret https"
            httpReq.Open "post", "https://" + m_ip + "/" + url, True
        End If
        DebugLog "[AJAXConnection] ReadyState " + CStr(httpReq.readyState) + " RefId:" + m_ref + "url:" + url + " send: " + data
    
        If m_ref <> Empty Then
            httpReq.setRequestHeader "Security-Hint", m_ref
        Else
            If m_tempHint <> Empty Then
                httpReq.setRequestHeader "Security-Hint", m_tempHint
            End If
        End If
        
        'httpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
      '//good
     'Set m_httpReq = CreateObject("MSXML2.XMLHTTP.6.0")
    'm_httpReq.Open "get", "https://" + m_ip + "/" + url, True
    'm_httpReq.Open "get", "https://192.168.0.53/", True
    'm_httpReq.Send
        
        httpReq.Send data
    End If
End Sub

Public Function GetAvailableHttpRequest()
    Dim i
    For i = LBound(m_operatingReq) To UBound(m_operatingReq) Step 1
        If m_operatingReq(i) Is Nothing Then
            Set m_operatingReq(i) = New HttpRequest
            m_operatingReq(i).SetId (i)
            Exit For
        Else
            If m_operatingReq(i).IsAvailable() Then
                Exit For
            End If
        End If
    Next
    
    If i <= UBound(m_operatingReq) Then
        Set GetAvailableHttpRequest = m_operatingReq(i)
    Else
        Set GetAvailableHttpRequest = Nothing
    End If
End Function

' The following instance might be called simultaneously
' 1. Call Back Handler > HttpRequest > Instantly Schedule > StartReq
' 2. Timeout > Periodical Schedule > StartReq
' 3. UI operation > HttpRequest > Instantly Schedule > StartReq
Public Sub StartReq()
    Do While Not IsQueueEmpty()
        Dim httpReq As HttpRequest
        
        ' Remarks: The following operation might trigger reschedule -- Consume Req
        ' However, the available request will not be consumed during operation because
        ' it is locked during reschedule.
        
        DebugLog "[AJAXConnection] Schedule"
        
        Set httpReq = GetAvailableHttpRequest()
        
        If httpReq Is Nothing Then
            DebugLog "[AJAXConnection] Schedule Quit"
            Exit Do
        Else
            Dim req
            Set req = DequeueReq()
            
            ' Double Check req Because it might be called simultaneously.
            ' Remarks: It could also be Nothing because GetAvailableHttpRequest
            ' might have been triggered reschedule already.
            If Not req Is Nothing Then
                req.InitiateRequest httpReq
            End If
        End If
        DebugLog "[AJAXConnection,StartReq] Schedule End"
    Loop
End Sub

Public Sub TryStartTimeout()
    If Not m_refreshing Then
        m_refreshing = True
        
        RegisterHeartBeatTimer "TryTimeout"
    End If
End Sub

Public Sub TryStopTimeout()
    If m_refreshing Then
        m_refreshing = False
        
        UnRegisterHeartBeatTimer "TryTimeout"
    End If
End Sub

Public Sub TryTimeout()
    Dim hasAvailable As Boolean
    Dim hasWorking As Boolean
    
    DebugLog "[AJAXConnection] Tick"
    
    Dim i
    For i = LBound(m_operatingReq) To UBound(m_operatingReq) Step 1
        If Not m_operatingReq(i) Is Nothing Then
            'DebugLog "[AJAXConnection] " + CStr(i) + " ReadyState " + CStr(m_operatingReq(i).GetHttp.readyState)
            
            ' tick each http request
            m_operatingReq(i).TickExecutionTime
        
            If m_operatingReq(i).IsAvailable Then ' call this function to trigger handling / timeout
                hasAvailable = True
            Else
                hasWorking = True
            End If
        End If
    Next
    
    If Not IsQueueEmpty() Then
        If hasAvailable Then
            StartReq ' Try start new request if it has some request available
        End If
    Else
        ' When no pending request nor working request, simply stop timeout
        If Not hasWorking Then
            TryStopTimeout
        End If
    End If
End Sub

