VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AJAXReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim m_url
Dim m_data
Dim m_errorAction
Dim m_msgAction
Dim m_arg0
Dim m_pNext As AJAXReq
Dim m_httpReq
Dim m_operationId
Dim m_expireDate As Date
Dim m_retry As Integer

Sub Initialize(url, data, errorAction, msgAction, arg0)
    m_url = url
    m_data = data
    m_errorAction = errorAction
    m_msgAction = msgAction
    m_arg0 = arg0
End Sub

Sub SetNext(pNext)
    Set m_pNext = pNext
End Sub

Function GetNext()
    Set GetNext = m_pNext
End Function

Sub SetRetry(retry)
    m_retry = retry
End Sub

Function GetRetry()
    GetRetry = m_retry
End Function

Sub InitiateRequest(httpReq As HttpRequest)
    'MsgBox "InitiateRequeste"
    httpReq.Initialize m_url, m_data, m_errorAction, m_msgAction, m_arg0, m_retry
End Sub
' no use , resurce is in httprequest
Sub StartRequest()
    Dim m_httpReq
    Set m_httpReq = CreateObject("MSXML2.XMLHTTP.6.0")   'CreateObject("Msxml2.ServerXMLHTTP.6.0")
    ' m_httpReq.setOption 2, SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
    'm_httpReq.objHTTP.SetOption(2) = (m_httpReq.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
    'm_httpReq.SetOption(2, m_httpReq.GetOption(2) - SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS)
    
    
    Dim req
    Set req = New HttpMonitor
    
    req.Initialize m_httpReq, m_errorAction, m_msgAction, m_arg0
    
    m_httpReq.onreadystatechange = req

    'm_httpReq.onreadystatechange = Me
    
    DoHttpRequest m_httpReq, m_url, m_data
End Sub

Public Sub Cancel()
    DebugLog "[AJAXReq] Cancelled"
    RunApp m_errorAction, m_arg0, 0 ' unspecified status code
End Sub

Public Sub RunApp(action, arg0, arg1)
    If action <> "" Then
        Application.Run action, arg0, arg1
    End If
End Sub

Private Sub Class_Terminate()
    'MsgBox "AJAXReq Terminate"
End Sub
