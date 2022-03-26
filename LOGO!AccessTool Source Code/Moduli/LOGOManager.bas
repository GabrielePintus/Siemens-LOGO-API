Attribute VB_Name = "LOGOManager"
Option Private Module
Public Sub InitializeLOGOManager()
    AddPropertyListener PROPERTY_ID_CONNECTION, "LOGOConnectionHandler"
End Sub

Public Sub LOGOConnectionHandler(connectionState, code)
    If connectionState = STATE_CONNECTED Then
        RegisterHeartBeatTimer "LOGOPeriodicalTask"
        RegisterHeartBeatTimer "CallTrendProcedure"
    Else
        UnRegisterHeartBeatTimer "LOGOPeriodicalTask"
        UnRegisterHeartBeatTimer "CallTrendProcedure"
        StopDataLog
    End If
End Sub

Public Sub LOGOPeriodicalTask()
    If GetProperty(PROPERTY_ID_CONNECTION) = STATE_CONNECTED Then
        RefreshAndMakeDataLog
    End If
End Sub

Private Sub RefreshAndMakeDataLog()
    UpdateHierarchy ' ensure real-time hierarchy
    
    ' create time stamp
    Dim logTimeStr As String
    logTimeStr = CStr(Format(Now, "yyyy-mm-dd hh:mm:ss"))
    
    ' Traverse and refresh all trendcharts
    Dim containers
    Set containers = GetWorkBookContainers()
    
    ' Update Variables
    TryUpdateVariables
    
    ' refresh all charts
    Dim entries
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        entries(i).StartVARL
    Next
    
    ' Trigger a update and collect log at the same time
    RefreshVolatile
    
    ' Make them record their own logs
    For i = 0 To containers.Count - 1 Step 1
        entries(i).CloseVARL logTimeStr
    Next
End Sub

Private Sub StopDataLog()
    UpdateHierarchy ' ensure real-time hierarchy
    
    ' Traverse and refresh all trendcharts
    Dim containers
    Set containers = GetWorkBookContainers()
    
    ' refresh all charts
    Dim entries
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        entries(i).StopRefreshing
    Next
End Sub
