Attribute VB_Name = "DataLogGadgetManager"
Option Private Module
Dim m_oRegExp As Object

Public Const GADGET_TYPE_DATALOG As String = "DataLog"

Public Sub RefreshAllDataLogs()
    ' create time stamp
    Dim timeStr As String
    Dim logTimeStr As String
    timeStr = CStr(Format(Time, "hh:mm:ss"))
    logTimeStr = CStr(Format(Now, "mm/dd/yy hh:mm:ss"))
    
    ' Traverse and refresh all trendcharts
    Dim containers
    Set containers = GetWorkBookContainers()
    
    ' refresh all charts
    Dim entries
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        Dim gadgetManager
        Set gadgetManager = entries(i).GetGadgetManager(GADGET_TYPE_DATALOG)
        
        If Not gadgetManager Is Nothing Then
            Dim gadgets
            Set gadgets = gadgetManager.GetGadgets()
        
            Dim gadgetEntries
            gadgetEntries = gadgets.items
            
            Dim j
            For j = 0 To gadgets.Count - 1
                gadgetEntries(j).RefreshData timeStr, logTimeStr
            Next
        End If
    Next
End Sub

Private Sub StopAllDataLogs()
    ' Traverse and stop all trendcharts
    Dim containers
    Set containers = GetWorkBookContainers()
    
    ' refresh all charts
    Dim entries
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        Dim gadgetManager
        Set gadgetManager = entries(i).GetGadgetManager(GADGET_TYPE_DATALOG)
        
        If Not gadgetManager Is Nothing Then
            Dim gadgets
            Set gadgets = gadgetManager.GetGadgets()
        
            Dim gadgetEntries
            gadgetEntries = gadgets.items
            
            Dim j
            For j = 0 To gadgets.Count - 1
                gadgetEntries(j).StopRefreshing
            Next
        End If
    Next
End Sub

Public Sub InitializeDataLogManager()
    ' Start Check LogIn Conflict
    RegisterHeartBeatTimer "DataLogPeriodicalTask"
    
    AddPropertyListener PROPERTY_ID_CONNECTION, "DataLogConnectionHandler"
End Sub

Public Sub DataLogConnectionHandler(connectionState, code)
    If connectionState <> STATE_CONNECTED Then
        ' Stop all trend charts when it is not in connected state
        StopAllDataLogs
    End If
End Sub

Public Sub DataLogPeriodicalTask()
    If GetProperty(PROPERTY_ID_CONNECTION) = STATE_CONNECTED Then
        RefreshAllDataLogs
    End If
End Sub

Public Function IsDataLogFileName(name)
    IsDataLogFileName = False ' return false by default to indicate error
    
    ' name consists of 14 digits and .csv
    If Len(name) <> 18 Then
        Exit Function
    End If

    If m_oRegExp Is Nothing Then
        Set m_oRegExp = CreateObject("vbscript.regexp")
        m_oRegExp.Pattern = "^\d+\.csv$"
    End If
    
    Dim omatches
    
    Set omatches = m_oRegExp.Execute(LCase(name))
    If omatches.Count = 1 Then
        IsDataLogFileName = True
    End If
End Function

