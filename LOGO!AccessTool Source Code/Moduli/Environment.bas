Attribute VB_Name = "Environment"
Option Private Module
Dim LOG
Dim fso As Object
Dim mutex As Boolean

' Remarks: HeartBeat Timer is more frequently referred than Environment. That's
' Why we let HeartBeatTimer initialize environment instead of the reverse.

' This interface will starts the process after calling
Public Sub InitializeEnvironment()
    ' Initialize LoginGadget Mangager
    'InitializeLogInGadgetManager

    ' Initialize Trend Chart Manager
    'InitializeTrendChartManager

    InitializeVariableSyncManager
    
    ' Initialize DataLog Manager
    'InitializeDataLogManager
    
    InitializeLOGOManager
    
    GetWorkBookContainers
End Sub

Public Sub ValidateEnvironment()
    UpdateHierarchy ' ensure real-time hierarchy
End Sub

Public Function GetLanguage()
    GetLanguage = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    'GetLanguage = 1031
End Function


Public Function GetFileSystemObject()
    If fso Is Nothing Then
        Set fso = CreateObject("Scripting.FileSystemObject")
    End If
    
    Set GetFileSystemObject = fso
End Function

Public Sub EnterCriticalSection()
    If mutex Then
        DebugLog "CriticalSection Conflicts Detected"
        While mutex
        Wend
        DebugLog "CriticalSection Conflicts Relief"
    End If
    
    mutex = True
End Sub

Public Sub LeaveCriticalSection()
    mutex = False
End Sub

'hid for debug, Do not release
Public Sub DebugLog(STR)
    'Exit Sub
    
    If TypeName(LOG) = "Empty" Then
        Set LOG = New DataLog
        LOG.OpenDir ("D:\log")
        LOG.SetHead ("====")
    End If
    
    Dim timeStr As String
    timeStr = CStr(Format(Now, "yyyy-mm-dd hh:mm:ss"))
    
    LOG.WriteLog (timeStr + " " + STR)
End Sub
