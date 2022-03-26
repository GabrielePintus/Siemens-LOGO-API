Attribute VB_Name = "ToolConfigManager"
Option Private Module
Public m_StopFlag As Integer  'from the UI
 Public m_StopFlagChange As Integer

 Public m_Interval As Integer 'from the UI
 Public m_DirNum As Integer
 Public m_OldDirNum As Integer
 Public m_OldInterval As Integer
Public m_FirstConnectFlag  As Integer

 
'可以是三个常数之一： ForReading 、 ForWriting 或 ForAppending
'分别是 1 ，2 ，8
 

 Public Function setConnectFlag(Interval As Integer)
    m_FirstConnectFlag = Interval
End Function

Public Function getConnectFlag()
    getConnectFlag = m_FirstConnectFlag
End Function
 
 Public Function setInterval(Interval As Integer)
    m_Interval = Interval
End Function

Public Function getInterval()
    getInterval = m_Interval
End Function


Public Function setStopFlag(stopFlag As Integer)
    m_StopFlag = stopFlag
End Function

Public Function getStopFlag()
    getStopFlag = m_StopFlag
End Function


Public Function SetStopFlagChange(temp As Integer)
    m_StopFlagChange = temp
End Function

Public Function GetStopFlagChange()
    GetStopFlagChange = m_StopFlagChange
End Function

Public Function GetIntervalChange()
    If m_OldInterval <> m_Interval Then
    m_OldInterval = m_Interval
    GetIntervalChange = 1
    Else
    GetIntervalChange = 0
    End If

End Function

Public Sub SynchronizeToFile()
    Dim fso As Object, configFileHandle As Object, configFilePath As String
    Dim configIP As String
    Set fso = GetFileSystemObject
    configFilePath = Environ("USERPROFILE") + "\Documents" + "\logoaccessconfigpath.txt"
    'blnExist = fso.FileExists(configFilePath)
    
    Set configFileHandle = fso.OpenTextFile(configFilePath, 2, True)
    configFileHandle.WriteLine (GetUrl)
    configFileHandle.WriteLine (m_Interval)
    configFileHandle.WriteLine (m_DirNum)
    
    configFileHandle.Close
End Sub
'only when Open
Public Sub SynchronizeFromFile()
    Dim fso As Object, configFilePath As String, configFileHandle As Object, blnExist As Boolean
    Dim configIP As String
    Set fso = GetFileSystemObject
    configFilePath = Environ("USERPROFILE") + "\Documents" + "\logoaccessconfigpath.txt"
    blnExist = fso.FileExists(configFilePath)
    If blnExist Then
        Set configFileHandle = fso.OpenTextFile(configFilePath, 1)
        configIP = configFileHandle.ReadLine
        m_Interval = configFileHandle.ReadLine
        m_DirNum = configFileHandle.ReadLine
        m_OldDirNum = m_DirNum
        SetUrl (configIP)
        configFileHandle.Close
        If (m_DirNum < 1) Or (m_DirNum > 100) Then
            m_DirNum = 5
        End If
        If (m_Interval < 1) Or (m_Interval > 60) Then
             m_Interval = 1
        End If
        m_OldDirNum = m_DirNum
        
    Else
    'Set configFileHandle = fso.OpenTextFile(ForAppending) create when logoin or changeconfig
    
    m_DirNum = 5
    m_OldDirNum = m_DirNum
    m_Interval = 1
    SetUrl ("192.168.0.3")
    End If

End Sub

