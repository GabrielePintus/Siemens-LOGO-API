Attribute VB_Name = "CallRangeContainerManager"
Option Private Module
Dim m_CallRangeContainers
Dim m_LogoTrenduid As Integer 'Uid for every trend
Dim m_AddCallStatus As Integer
Dim m_AddCallNum As Integer  'The formula num user added.
Private Const MAX_INVALID_RANGE_NUM As Integer = 10
Dim m_invalidFormulaRangeArray(MAX_INVALID_RANGE_NUM) As Range

Public Const SCAN_FINISHED As Integer = 2
Public Const INCREASE_FORMULA As Integer = 1
Public Const INCREASE_FINISHED As Integer = 0
Public Function GetCallRangeContainers()
    'MsgBox "Called"
    
    If TypeName(m_CallRangeContainers) = "Empty" Then
        Set m_CallRangeContainers = CreateObject("Scripting.Dictionary")
    End If
    
    Set GetCallRangeContainers = m_CallRangeContainers
End Function

Public Function CallTrendProcedure()
  On Error Resume Next
  
  RemoveCallRangeContainer
  RemoveInvalidFormulaRangeArray
  
    Dim keys
    Dim entries
     Dim sngBegin As Single
     Dim sngEnd As Single
     
     Dim OneBegin As Single
     Dim OneEnd As Single
    Dim containers
    Set containers = GetCallRangeContainers()
    
    keys = containers.keys
    entries = containers.items
   Dim i As Integer
            
    DebugLog "[CallTrendProcedure] Start"
  ' only for syn the formula
  '需要增加公式的三种情况：添加公式， 第一次运行已经有的公式，增加长度(m_OldDirNum < m_DirNum)
    If m_AddCallStatus = SCAN_FINISHED Or m_OldDirNum < m_DirNum Then
        DebugLog "[CallTrendProcedure] add formula"
         If m_OldDirNum < m_DirNum Then
                For i = 0 To containers.Count - 1 Step 1
                    If Not entries(i) Is Nothing Then
                        entries(i).SetFormulaSynFLag 'All the Formula should be set again
                    End If
                Next
             m_OldDirNum = m_DirNum
        End If
   '创建进度条
       ' prgramBarShow.Show
       ' Appolication.ScreenUpdating = False
      
        'Dim tempIndex As Integer
        
       ' prgramBarShow.Repaint
          
        'sngBegin = Timer
        'Dim entry
        '添加公式
        If m_AddCallNum = 1 Then
            entries(containers.Count - 1).AddFormula 1, 0
        Else ' 第一次运行已经有的公式，增加长度(m_OldDirNum < m_DirNum)
        
        For i = 0 To containers.Count - 1 Step 1
            entries(i).AddFormula containers.Count, i
          
           '按任务完成情况的百分比设置lblProgress的宽度，实现模拟进度效果
       ' prgramBarShow.lblprogress.Width = prgramBarShow.lblBack.Width * i / containers.Count
            '格式化显示完成任务情况的百分比
      
       ' prgramBarShow.percert = Format(i / containers.Count, "0%")
            '将进度自身重绘，实时显示
       ' prgramBarShow.Repaint
        Next
        
        End If
        m_AddCallStatus = INCREASE_FINISHED 'syn finished,wait for next AddCallRange
        m_AddCallNum = 0
        'pbar.DestroyBar

         'Set pbar = Nothing
  
       Unload prgramBarShow
       
       'prgramBarShow.Hide
       
        End If  'm_AddCallStatus = SCAN_FINISHED
            
           
        'End If
       ' Next
        'sngEnd = Timer
        'Debug.Print "Total:"; sngEnd - sngBegin
         If GetTrendFlush = 1 Then
            For i = 0 To containers.Count - 1 Step 1
             
               If Not entries(i) Is Nothing Then
               entries(i).UpdateHisdata
               End If
            Next
         SetTrendFlush (0) 'The his trend data update is also influnced by the m_Interval and Stopflag
         End If
        If m_OldDirNum > m_DirNum Then
             For i = 0 To containers.Count - 1 Step 1
          
                If Not entries(i) Is Nothing Then
                entries(i).DecreaceDirNum
                End If
             Next
             
      
            m_OldDirNum = m_DirNum
        End If
        'valid call ranges
        Dim isValid As Boolean
        isValid = True
        For i = 0 To containers.Count - 1 Step 1
                If Not entries(i) Is Nothing Then
                    isValid = entries(i).ValidCallRange
                    If Not isValid Then
                        m_AddCallStatus = SCAN_FINISHED
                        Exit For
                    End If
                End If
        Next
        
  '  End If

DebugLog "[CallTrendProcedure] End"
End Function

'CallRange As Object, DirType As Integer, DirNum As Integer
Public Function AddCallRange(DirType As Integer, DirNum As Integer, CallRange As Range)
 Dim findRet As Integer
 Dim keys
 Dim entries
 Dim containers
 Set containers = GetCallRangeContainers()
 Dim tempCallRangeContainer As CallRangeContainer
 
 
 If GetProperty(PROPERTY_ID_CONNECTION) <> STATE_CONNECTED Then
    Exit Function
 End If
 
 
 'CallRangeContainer tempCallRangeContainer
  entries = containers.items
' keys = containers.keys
 
    findRet = FindCallRange(CallRange, DirType)
    
      'Is tianjia
   If findRet <> 0 Then
      Set tempCallRangeContainer = entries(findRet - 1)
        tempCallRangeContainer.UpdateExpireDate
        If m_AddCallStatus = INCREASE_FORMULA Then
          m_AddCallStatus = SCAN_FINISHED ' time to syn the formula
        End If
   Else  'If the one HisTrend is added, The dialog is shown, until the formula is added.
        'If containers.count = 0,
        Set tempCallRangeContainer = New CallRangeContainer
        tempCallRangeContainer.Initialize m_LogoTrenduid, DirType, DirNum, CallRange
        If Not containers.exists(m_LogoTrenduid) Then
            containers.Add m_LogoTrenduid, tempCallRangeContainer
            m_LogoTrenduid = m_LogoTrenduid + 1
            m_AddCallStatus = INCREASE_FORMULA ' Increasing
            m_AddCallNum = m_AddCallNum + 1
        Else
            containers.item(m_LogoTrenduid) = tempCallRangeContainer
        End If
     End If
End Function
' return bool
'tfs2525512
Public Function FindCallRange(CallRange As Range, DirType As Integer)
     Dim containers
     Dim items
    Set containers = GetCallRangeContainers()
    items = containers.items
    Dim tempCallRangeContainer As CallRangeContainer
    'CallRangeContainer tempCallRangeContainer
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        Dim key
        Set tempCallRangeContainer = items(i)
        
        If tempCallRangeContainer.m_Column = CallRange.Column And tempCallRangeContainer.m_Row = CallRange.Row _
        And tempCallRangeContainer.m_WorkSheet Is CallRange.Worksheet And tempCallRangeContainer.m_DirType = DirType Then
            FindCallRange = 1 + i
            Exit Function
        End If
    Next
    FindCallRange = 0
End Function

Public Function FindCallRangeContainerByUID(LogoTrenduid As Integer)
     Dim containers
     Dim keys
    Set containers = GetCallRangeContainers()
    keys = containers.keys
    Dim tempCallRangeContainer As CallRangeContainer
    'CallRangeContainer tempCallRangeContainer
    If containers.exists(LogoTrenduid) Then
        Set FindCallRangeContainerByUID = containers.item(LogoTrenduid)
        Exit Function
    End If
    FindCallRangeContainerByUID = Empty
End Function
Private Sub RemoveInvalidFormulaRangeArray()
    For i = 0 To MAX_INVALID_RANGE_NUM Step 1
        If Not m_invalidFormulaRangeArray(i) Is Nothing Then
            m_invalidFormulaRangeArray(i).ClearContents
            m_invalidFormulaRangeArray(i) = Nothing
        End If
    Next
End Sub
'this should be put into the 1s cycle task
Public Sub RemoveCallRangeContainer()
    Dim keys
    Dim entries
    
    Dim containers
    Set containers = GetCallRangeContainers()
    
    keys = containers.keys
    entries = containers.items
    
    Dim i
    For i = 0 To containers.Count - 1 Step 1
        Dim entry
        Set entry = entries(i)
        If Not entry Is Nothing Then
            entries(i).TickIdleTime ' Tick to increment the idle time
    
            ' remove entry if it has been expired
            If entries(i).IsExpired() Then
                entries(i).RemoveTrendFormula
                containers.Remove (keys(i))
                'Erase keys(i)
             
            End If
        End If
    Next
End Sub
 Public Function GetTrendHisValue(TrendUID As Integer, TrendIndex As Integer)
    On Error GoTo Err
    Application.Volatile
    Dim tempCallRangeContainer As CallRangeContainer
    Dim typevalue

    If Not IsEmpty(FindCallRangeContainerByUID(TrendUID)) Then
        Set tempCallRangeContainer = FindCallRangeContainerByUID(TrendUID)
        typevalue = tempCallRangeContainer.GetValueByIndex(TrendIndex)
        GetTrendHisValue = typevalue
    Else
        Dim trend_cell As Range
        If TypeName(Application.Caller) = "Range" Then
             Set trend_cell = Application.Caller
             For i = 0 To MAX_INVALID_RANGE_NUM Step 1
                If m_invalidFormulaRangeArray(i) Is Nothing Then
                    Set m_invalidFormulaRangeArray(m_invalidFormulaNum) = trend_cell
                    Exit For
                End If
             Next
             'trend_cell.ClearContents 'Cannot clear formula when it's in running state, to clear it, record it first and clear it in cycle run
        End If
      GetTrendHisValue = Empty
    End If
    Exit Function
Err:
        GetTrendHisValue = "Invalid VAR Log" ' return invalid VAR Log to indicate error

 End Function
 
  Public Function Test123(TrendUID As Integer)
 On Error GoTo Err
    Application.Volatile
    ' Dim tempCallRangeContainer As CallRangeContainer
   ' GetTrendHisValue = TrendUID + TrendIndex
   Test123 = TrendUID
    Exit Function
Err:
        Test123 = "Invalid VAR Log" ' return invalid VAR Log  to indicate error

 End Function

