Attribute VB_Name = "TrendChartManager"
Option Private Module
Public Const GADGET_TYPE_TRENDCHART As String = "TrendChart"

Public Sub RefreshAllTrendCharts()
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
        Set gadgetManager = entries(i).GetGadgetManager(GADGET_TYPE_TRENDCHART)
        
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

Public Sub InitializeTrendChartManager()
    ' Start Check LogIn Conflict
    RegisterHeartBeatTimer "TrendChartPeriodicalTask"
End Sub

Public Sub TrendChartPeriodicalTask()
    If GetProperty(PROPERTY_ID_CONNECTION) = STATE_CONNECTED Then
        RefreshAllTrendCharts
    End If
End Sub
