Attribute VB_Name = "LogInGadgetManager"
Option Private Module
'total no use by chen
Dim m_oIpValidateRegExp
Dim m_lastLogInWorkbook As Workbook ' remember which one logged in
Dim m_lastLogInGadget

Public Const LOGIN_GADGET_STATE_NORMAL As Integer = 0
Public Const LOGIN_GADGET_STATE_CONNECTING As Integer = 1
Public Const LOGIN_GADGET_STATE_CONNECTED As Integer = 2
Public Const LOGIN_GADGET_STATE_DISABLED As Integer = 3

Public Const GADGET_TYPE_LOGIN As String = "LogIn"

Public Function GetIPValidateReg(ipAddr)
    If TypeName(m_oIpValidateRegExp) = "Empty" Then
        Set m_oIpValidateRegExp = CreateObject("vbscript.regexp")
        m_oIpValidateRegExp.Pattern = "^(\d+)\.(\d+)\.(\d+)\.(\d+)$"
    End If
    
    Dim omatches
    
    Set omatches = m_oIpValidateRegExp.Execute(ipAddr)
    
    If omatches.Count = 1 Then
        Dim i
        For i = 0 To 3
            Dim valueStr
            Dim value
            
            valueStr = omatches.item(0).SubMatches.item(i)
            
            ' Since VBScript Regular Expression doesn't support {} operator well, this is a workaround
            ' to avoid overflow of CInt() function which doesn't support large integer
            If Len(valueStr) > 3 Then
                GetIPValidateReg = False
                Exit Function
            End If
            
            value = CInt(valueStr)
            
            If value > 255 Then
                GetIPValidateReg = False
                Exit Function
            End If
        Next
        
        GetIPValidateReg = True
    Else
        GetIPValidateReg = False
    End If
End Function

Public Function FormatIP(ipAddr)
    If TypeName(m_oIpValidateRegExp) = "Empty" Then
        Set m_oIpValidateRegExp = CreateObject("vbscript.regexp")
        m_oIpValidateRegExp.Pattern = "^(\d+)\.(\d+)\.(\d+)\.(\d+)$"
    End If
    
    Dim omatches
    
    Set omatches = m_oIpValidateRegExp.Execute(ipAddr)
    Dim formattedIp
    
    If omatches.Count = 1 Then
        Dim i
        For i = 0 To 3
            Dim valueStr
            Dim value
            
            valueStr = omatches.item(0).SubMatches.item(i)
            
            ' Since VBScript Regular Expression doesn't support {} operator well, this is a workaround
            ' to avoid overflow of CInt() function which doesn't support large integer
            If Len(valueStr) > 3 Then
                FormatIP = ""
                Exit Function
            End If
            
            value = CInt(valueStr)
            
            If value > 255 Then
                FormatIP = ""
                Exit Function
            End If
            
            ' Append the new section
            If formattedIp <> "" Then
                formattedIp = formattedIp + "."
            End If
            formattedIp = formattedIp + CStr(value)
        Next
        
        FormatIP = formattedIp
    Else
        FormatIP = ""
    End If
End Function

Public Sub GadgetLogOut(gadgetId)
    ' HandleAllLogInGadgets LOGIN_GADGET_STATE_NORMAL ' reset all gadgets to 0 state
    
    LogOut
End Sub

Public Sub GadgetLogIn(ip, pwd, gadgetId)
    ' The ip, pwd is assumed to be valid
    
    ' TODO: logout if necessary
    
    Set m_lastLogInWorkbook = Application.ActiveWorkbook
    m_lastLogInGadget = gadgetId
    
    LogIn ip, pwd, "OnLogInFail", "OnLogInSucess", gadgetId ' pass gadgetId as arg0
    
    ' disable all gadgets, the operating gadget enters login status
    UpdateAllExcept LOGIN_GADGET_STATE_DISABLED, m_lastLogInWorkbook, gadgetId, LOGIN_GADGET_STATE_CONNECTING
End Sub

Public Sub OnLogInSucess(arg0, status)
    'MsgBox "Log in successfully" ' Now() - cur  magic + CStr(status)

    ' enable all gadgets, the operating gadget enters connected state
    UpdateAllExcept LOGIN_GADGET_STATE_NORMAL, m_lastLogInWorkbook, arg0, LOGIN_GADGET_STATE_CONNECTED
End Sub

Public Sub OnLogInFail(arg0, status)
    MsgBox "Failed to Log In"
    ' Don't need to update gadgets because it will be updated by connection handler
    'HandleAllLogInGadgets LOGIN_GADGET_STATE_NORMAL ' enable all gadgets
End Sub

Public Sub InitializeLogInGadgetManager()
    ' Start Check LogIn Conflict
    'RegisterHeartBeatTimer "LogInGadgetPeriodicalTask"
    
    ' It shall be unregistered when document being closed
    AddPropertyListener PROPERTY_ID_CONNECTION, "LoginGadgetManagerConnectionHandler"
    
    AddPropertyListener PROPERTY_ID_STATUS, "LoginGadgetManagerStatusHandler"
End Sub

Public Sub LogInGadgetPeriodicalTask()
    
End Sub

Public Sub LoginGadgetManagerConnectionHandler(connectionState, code)
    If connectionState = STATE_NOT_CONNECTED Then
        HandleAllLogInGadgets LOGIN_GADGET_STATE_NORMAL ' enable all gadgets
    End If
End Sub

Public Sub LoginGadgetManagerStatusHandler(newValue, arg0)
    Dim LogInGadget
    Set LogInGadget = GetGadget(m_lastLogInWorkbook, m_lastLogInGadget)
    
    If Not LogInGadget Is Nothing Then
        LogInGadget.SetStatus newValue
    End If
End Sub

Private Sub HandleAllLogInGadgets(style)
    Dim containers
    Set containers = GetWorkBookContainers()
    
    ' refresh all charts
    Dim entries
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        Dim gadgetManager
        Set gadgetManager = entries(i).GetGadgetManager(GADGET_TYPE_LOGIN)
        
        If Not gadgetManager Is Nothing Then
            Dim gadgets
            Set gadgets = gadgetManager.GetGadgets()
        
            Dim gadgetEntries
            gadgetEntries = gadgets.items
            
            Dim j
            For j = 0 To gadgets.Count - 1
                gadgetEntries(j).SetStyle style
            Next
        End If
    Next
End Sub

Private Sub UpdateAllExcept(allStatus, Workbook As Workbook, exceptGadgetId, exceptStatus)
    HandleAllLogInGadgets allStatus
    
    ' get workbook container
    Dim LogInGadget
    Set LogInGadget = GetGadget(Workbook, exceptGadgetId)
    
    If LogInGadget Is Nothing Then
        Exit Sub
    End If

    LogInGadget.SetStyle exceptStatus
End Sub

Public Function GetGadget(Workbook As Workbook, gadgetId)
    Set GetGadget = Nothing ' return nothing by default to indicate error

    Dim workbookContainer
    Set workbookContainer = GetWorkBookContainer(Workbook)
    
    If workbookContainer Is Nothing Then
        Exit Function
    End If
    
    Dim LogInGadget
    Set LogInGadget = workbookContainer.GetGadget(GADGET_TYPE_LOGIN, gadgetId)
    
    Set GetGadget = LogInGadget
End Function
