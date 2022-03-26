Attribute VB_Name = "VariableSyncManager"
Option Private Module
Dim m_variables
Public lngTimerID As Long
Public m_onGoingRequestCount As Integer
Dim m_oRegExp
Dim m_variableFactory As Object
Dim m_array As SortArray
Dim m_packer As Object
Dim m_data As Object
Dim datalogrecord As Integer
Dim trendflushflag As Integer


'Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Public Const VAR_SIZE_BIT = 1
Public Const VAR_SIZE_BYTE = 2
Public Const VAR_SIZE_WORD = 4
Public Const VAR_SIZE_DWORD = 6

Public Const MAX_SIZE_REQUEST = 1024
Public Const MAX_SIZE_RESPONSE = 1015
Public Const MAX_SIZE_GAP = 10

Public Const MAX_VARIABLE_EACH_REQUEST = 25
' 1024(without head) to contain items atmost 40 bytes <r i='VMBA1024' e='-255' v='0000ffff'/>

Public Sub RefreshVolatile()
    ' Always trigger for refresh now, even it will manually call Calculate to Refresh Workbook by Workbook
    ' Because that feature is bound with LOG
    Sheet1.Cells(1, 1) = ""
End Sub

Public Sub RemovePacker()
    Set m_packer = Nothing
End Sub

Public Function GetPacker()
    If Not m_packer Is Nothing Then
        Set GetPacker = m_packer
        Exit Function
    End If

    Dim variables
    Set variables = GetVariables()
    keys = variables.keys
    entries = variables.items
    
    Dim SortArray As SortArray
    Set SortArray = GetSortArray()
    
    ' Clear current variables
    SortArray.Clear
    
    ' Collect all variables
    Dim validCount As Integer
    Dim i
    For i = 0 To variables.Count - 1 Step 1
        Dim variable As Object
        Set variable = entries(i)
        If Not variable Is Nothing Then
            SortArray.Add "", 0, variable
            validCount = validCount + 1
        End If
    Next
    
    DebugLog "[VariableSyncProcedure] [" + CStr(validCount) + "/" + CStr(variables.Count) + "] Create Packer"
    
    ' Sort all variables
    SortArray.Sort
    
    ' Make requests according to range and gaps
    ' The overhead to carry out a response is 21 which means response size = 21 + 2 * bytes
    ' This means any gap which is larger than 10 bytes cannot be compensated by filling.
    
    Select Case GetConnectionType()
    Case CONNECTION_0BA8
        Set m_packer = New VariablePackerV8
    Case Else
        Set m_packer = New VariablePackerV81
    End Select
    
    For i = 0 To SortArray.Count() - 1 Step 1
        If Not m_packer.AddVariable(SortArray.GetItem(i).m_variableEntry) Then
            Exit For
        End If
    Next
    
    Set GetPacker = m_packer
End Function

Public Function GetVariables()
    ' Perform Lazy Initialization
    If TypeName(m_variables) = "Empty" Then
        Set m_variables = CreateObject("Scripting.Dictionary")
        
        'AddConnectionListener "VariableSyncerConnectionHandler"
    End If
    
    Set GetVariables = m_variables
End Function

Public Function GetSortArray()
    If m_array Is Nothing Then
        Set m_array = New SortArray
    End If
        
    Set GetSortArray = m_array
End Function

Public Function GetVariableRegExp()
    If TypeName(m_oRegExp) = "Empty" Then
        Set m_oRegExp = CreateObject("vbscript.regexp")
        m_oRegExp.Pattern = "^([a-zA-Z_]+)(\d+)(?:\.(\d+))?$"
    End If
    
    Set GetVariableRegExp = m_oRegExp
End Function

Public Sub InitializeVariableSyncManager()
    AddPropertyListener PROPERTY_ID_CONNECTION, "VariableSyncerConnectionHandler"
End Sub

Public Sub ClearVariables()
    GetVariables().RemoveAll
End Sub

Public Sub VariableSyncProcedure()
    On Error Resume Next
    Static synNum As Integer
    Static filterNum As Integer
    
    'RefreshVolatile ' trigger recalculate and register
    
    DebugLog "[VariableSyncProcedure],m_onGoingRequestCount: " + CStr(m_onGoingRequestCount)
    
    ' Remove expired entries before syncing
    RemoveExipreVariableEntries
    
    ' If there is still some requests running on this schedule, skip this schedule and try
    ' another schedule
    If m_onGoingRequestCount > 0 Then
        Exit Sub
    End If
    
    ' It is probably that the execution process been interrupted by user operation (Log Out)
    If GetProperty(PROPERTY_ID_CONNECTION) <> STATE_CONNECTED Then
        DebugLog "[VariableSyncProcedure] Already Closed"
        Exit Sub
    End If
    
    filterNum = filterNum + 1
    synNum = synNum + 1
    If filterNum >= m_Interval Or GetStopFlagChange = 1 Or GetIntervalChange = 1 Or getConnectFlag = 1 Then
       If getStopFlag = 0 Then ' not stop
       
        FilterGetVariable
      
        Setdatalogrecord (1)
        SetTrendFlush (1)
        filterNum = 0
        synNum = 0
        
        If GetStopFlagChange = 1 Then
        SetStopFlagChange (0)
        End If
        If getConnectFlag = 1 Then
        DebugLog "[VariableSyncProcedure ] getConnectFlag=1"
        setConnectFlag 0
        End If
       End If
    End If
    
         ' just used to keep alive, even if interval time is 60s
        Dim timeStr As String
        Dim logTimeStr As String
        
        If synNum >= 10 Then
        SendGetStatusRequest
         timeStr = CStr(Format(Time, "hh:mm:ss"))
        logTimeStr = CStr(Format(Now, "mm/dd/yy hh:mm:ss"))
        DebugLog "[Common] SendGetStatusRequest is called for keep live" + timeStr + logTimeStr
        synNum = 0
        End If
      
   DebugLog "[VariableSyncProcedure end]"
 
    
   
       
End Sub

Public Function Setdatalogrecord(temp As Integer)
    datalogrecord = temp
End Function

Public Function Getdatalogrecord()
    Getdatalogrecord = datalogrecord
End Function
'used for CallRangeContainerManager
Public Function SetTrendFlush(temp As Integer)
    trendflushflag = temp
End Function

Public Function GetTrendFlush()
    GetTrendFlush = trendflushflag
End Function



Private Sub FilterGetVariable()
  
  
    
    ' Get Request Str at very beginning in case it will interfere event handling of http request
    Dim requestStr As String
    
    Dim packer
    
    Set packer = GetPacker()
    requestStr = packer.GetRequestStr()
   
    ' Try sync status
    SendGetStatusRequest
    
    ' Pack variables and sync
    DebugLog "[FilterGetVariable->SendVariableSyncRequest]" + requestStr
    SendVariableSyncRequest requestStr
    
    
    
    Exit Sub 'old code, should be deleate
    
    Dim keys
    Dim entries

    Dim getVarReqData
    
    ' Remove expired entries before syncing
    RemoveExipreVariableEntries
    
    ' initialize req data
    getVarReqData = ""
    
    ' traverse to compose req
    Dim variables
    Set variables = GetVariables()
    keys = variables.keys
    entries = variables.items
    
    Dim countInReq As Integer
    
    Dim i
    For i = 0 To variables.Count - 1 Step 1
        If Not entries(i) Is Nothing Then
            If getVarReqData <> "" Then
                getVarReqData = getVarReqData + ";"
            End If
            
            ' Append the variable in request
            getVarReqData = getVarReqData + keys(i) + "," + entries(i).GetReqStr()
            countInReq = countInReq + 1
            
            ' If enough for a request, issue it now
            If countInReq >= MAX_VARIABLE_EACH_REQUEST Then
                SendVariableSyncRequest getVarReqData
                
                getVarReqData = ""
                countInReq = 0
            End If
        End If
    Next

    ' try to sync variables if necessary
    SendVariableSyncRequest getVarReqData
    
    DebugLog "[VariableSyncProcedure] End"
End Sub

Private Sub SendVariableSyncRequest(variableStr)
    ' Only send request when it is still in Connected state. This could happen because
    ' The process might be interrupted by asynchronous handler which may change something
    If variableStr <> "" And GetProperty(PROPERTY_ID_CONNECTION) = STATE_CONNECTED Then
        HttpTimeoutRequest "AJAX", "GETVARS:" + variableStr, "CommonOnConnectionError", "GetVarMsgHandler", "", 1
        m_onGoingRequestCount = m_onGoingRequestCount + 1
    End If
End Sub

Public Sub SendGetStatusRequest()
    HttpTimeoutRequest "AJAX", "GETSTDG", "CommonOnConnectionError", "GetStatusHandler", "", 1
    m_onGoingRequestCount = m_onGoingRequestCount + 1
End Sub

Private Sub GetStatusHandler(arg0, data As String)
    If m_onGoingRequestCount > 0 Then
        m_onGoingRequestCount = m_onGoingRequestCount - 1
    End If
    
    If data = "" Then
        ' TODO: Error happens
        Exit Sub
    End If
    DebugLog "[GetStatusHandler] Response: " + data + "m_onGoingRequestCount:" + CStr(m_onGoingRequestCount)
    
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    
    xmlDoc.async = False
    xmlDoc.LoadXML data
    
    If xmlDoc.parseError.ErrorCode = 0 Then
        Set xmlRoot = xmlDoc.getElementsByTagName("r")
        
        If xmlRoot.Length = 1 Then
            Dim status
            Select Case xmlRoot.item(0).Text
            Case "Running"
                status = 2
            Case Else
                status = 1
            End Select
        
            SetProperty PROPERTY_ID_STATUS, status, ""
            
            'RefreshVolatile
        End If
    End If
End Sub

Private Function ParseHexStr(STR)
    Dim val
    val = CLng("&H" + STR)
    If val < 0 Then
        val = 4294967296# + val
    End If
    ParseHexStr = val
End Function

Public Sub TryUpdateVariables()
    ' Update value of all entries
    If Not m_data Is Nothing Then
        ' Take a snap shot of data
        Dim data
        Set data = m_data
        Set m_data = Nothing

        Dim variables
        Set variables = GetVariables()
        entries = variables.items
        
        ' Handle all variables
        Dim i
        For i = 0 To variables.Count - 1 Step 1
            Dim variable As Object
            Set variable = entries(i)
            If Not variable Is Nothing Then
                variable.UpdateValue data.GetValue(variable)
            End If
        Next
    End If
End Sub

Private Sub GetVarMsgHandler(arg0, data As String)
    If m_onGoingRequestCount > 0 Then
        m_onGoingRequestCount = m_onGoingRequestCount - 1
    End If
    
    If data = "" Then
        Exit Sub
    End If
    DebugLog "[GetVarMsgHandler]m_onGoingRequestCount:" + CStr(m_onGoingRequestCount)
    DebugLog "[GetVarMsgHandler] Response: " + data
    
    Set xmlDoc = CreateObject("Microsoft.XMLDOM")
    
    xmlDoc.async = False
    xmlDoc.LoadXML data
    
    If xmlDoc.parseError.ErrorCode = 0 Then
    
        'MsgBox "GetVar Handling"
    
        Set xmlRoot = xmlDoc.getElementsByTagName("r")
        If Not m_packer Is Nothing Then
            Set m_data = m_packer.GetData(xmlRoot)
            
            ' Don't update here, let timer try to update before refresh
            'TryUpdateVariables
        End If
    End If
End Sub

Public Function GetVarValue(entry)
    If m_data Is Nothing Then
        GetVarValue = Empty
    End If
    
    GetVarValue = m_data.GetValue(entry)
End Function

Public Sub RemoveExipreVariableEntries()
    Dim keys
    Dim entries
    
    Dim variables
    Set variables = GetVariables()
    
    keys = variables.keys
    entries = variables.items
    
    Dim i
    For i = 0 To variables.Count - 1 Step 1
        Dim entry
        Set entry = entries(i)
        If Not entry Is Nothing Then
            entries(i).TickIdleTime ' Tick to increment the idle time
    
            ' remove entry if it has been expired
            If entries(i).IsExpired() Then
                variables.Remove (keys(i))
                
                ' Packer need to be updated
                RemovePacker
            End If
        End If
    Next
End Sub

Public Sub UpdateVariable(id, value)
    Dim variables
    Set variables = GetVariables()
    
    If variables.exists(id) Then
        variables.item(id).UpdateValue value
    End If
End Sub

Public Function GetCurTime()
   Dim logTimeStr As String
    GetCurTime = CStr(Format(Now, "yyyy-mm-dd hh:mm:ss"))
End Function

Public Function GetStatusString()
    Select Case GetProperty(PROPERTY_ID_CONNECTION)
    Case STATE_NOT_CONNECTED
        GetStatusString = "Offline"
    Case STATE_CONNECTING
        GetStatusString = "Connecting"
    Case STATE_CONNECTED
        Select Case GetProperty(PROPERTY_ID_STATUS)
        Case 1
            GetStatusString = "Stop"
        Case 2
            GetStatusString = "Running"
        Case Else ' Including Case 0
            GetStatusString = "Offline"
        End Select
    Case STATE_RECOVERING
        GetStatusString = "Recovering"
    Case Else
        GetStatusString = "InvalidStatus"
    End Select
End Function

Public Function GetVariableFactory()
    Set GetVariableFactory = m_variableFactory
End Function

Public Sub ClearVariableFactory()
    Set m_variableFactory = Nothing
End Sub

Public Function ValidateNumber(val, min, max)
    If val >= min And val <= max Then
        ValidateNumber = True
    Else
        ValidateNumber = False
    End If
End Function

Public Function GetVariableEntry(id, needAddInvalidOne As Boolean)
    Set GetVariableEntry = Nothing ' use nothing as default to indicate error
    
    ' If factory is undetermined, do nothing
    Dim factory
    Set factory = GetVariableFactory()
    
    If factory Is Nothing Then
        Exit Function
    End If

    ' format the id
    Dim formattedId
    formattedId = UCase(id)
    
    Dim entry
    Set entry = Nothing ' initialize it as nothing in case operation

    ' try to register new entry if it doesn't exist
    Dim variables
    Set variables = GetVariables()
    
    If Not variables.exists(formattedId) Then
        Set omatches = GetVariableRegExp().Execute(formattedId)
        If omatches.Count = 1 Then
            Dim Range As String
            Dim Addr As String
            Dim subAddr As String
            
            Range = omatches.item(0).SubMatches.item(0)
            Addr = omatches.item(0).SubMatches.item(1)
            subAddr = omatches.item(0).SubMatches.item(2)
            
            Set entry = CreateVariableEntry(factory, Range, Addr, subAddr)
        End If
        
        ' set formatted id as its req id
        If Not entry Is Nothing Then
            entry.SetReqId formattedId
            variables.Add formattedId, entry
            'make sure the syn from BM when new LOGOVAR added
            SetStopFlagChange (1)
            ' current variable packer need to be updated
            RemovePacker
        Else
            ' append the newly created entry, even if it is nothing (to avoid re-checking)
            If needAddInvalidOne Then
                variables.Add formattedId, entry
            End If
        End If
    Else
        Set entry = variables.item(formattedId)
    End If
    
    Set GetVariableEntry = entry
End Function

Public Function CreateVariableEntry(factory, Range, Addr, subAddr)
    On Error GoTo Err ' On Any error during attemption to create a variable entry, return nothing
    
    ' we shall specify rangeId, addr, type according to them
    If subAddr <> "" Then
        Set CreateVariableEntry = factory.Create2DVariableEntry(Range, CInt(Addr), CInt(subAddr))
    Else
        Set CreateVariableEntry = factory.CreateVariableEntry(Range, CInt(Addr))
    End If
    
    Exit Function
    
Err:
    Set CreateVariableEntry = Nothing ' return nothing for exception entry
End Function

Public Function GetVariableValue(id)
    Dim formattedId
    formattedId = UCase(id)
    
    Select Case formattedId
    Case "STATUS"
        GetVariableValue = GetStatusString()
    Case Else
        Dim entry
        Set entry = GetVariableEntry(id, True) ' When it is really asked to get value, add it even invalid
        
        If Not entry Is Nothing Then
            GetVariableValue = entry.GetValue()
        Else
            GetVariableValue = Empty ' this will make it invalid value. use it as default
        End If
    End Select
End Function

Public Sub VariableSyncerConnectionHandler(connectionState, code)
    If connectionState = STATE_CONNECTED Then
        'MsgBox "Syncer Connected"
        
        m_onGoingRequestCount = 0 ' Reset on going request count
        
        Select Case GetConnectionType()
        Case CONNECTION_0BA8
            Set m_variableFactory = New VariableFactoryV8
        Case Else
            Set m_variableFactory = New VariableFactoryV81
        End Select
        setConnectFlag 1
        RefreshVolatile ' Trigger each one to recalculate. This make a initial register of variable
        RegisterHeartBeatTimer "VariableSyncProcedure"
    Else
        'MsgBox "Syncer Not Connected"
        
        UnRegisterHeartBeatTimer "VariableSyncProcedure"
        
        DebugLog "[Module->VariableSynManager]Syncer Stop Running"
        
        ClearVariables
        ClearVariableFactory
        SetProperty PROPERTY_ID_STATUS, 0, "" ' Reset status to false when disconnected
        
        Set m_data = Nothing ' Clear data
        Set m_packer = Nothing ' Clear Packer
    End If
End Sub

Public Function NewVariableEntry(Range As Integer, Addr As Integer, size As Integer)
    ' Skip those invalid variables
    If Addr < 0 Or size <= 0 Then
        Set NewVariableEntry = Nothing
        Exit Function
    End If

    Dim entry
    
    Set entry = New VariableEntry
    entry.Initialize Range, Addr, size
    
    Set NewVariableEntry = entry
End Function
