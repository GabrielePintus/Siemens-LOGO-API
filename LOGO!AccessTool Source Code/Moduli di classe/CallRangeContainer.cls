VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CallRangeContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public m_WorkSheet As Worksheet

Public m_DirType As Integer
Public m_Row As Integer
Public m_Column As Integer

Public m_TrendNum As Integer ' maybe some new requirement will be this par should different
Public m_TrendChangeIndex As Integer  ' used to change the formula
Dim m_executionTime As Integer
Dim m_FormulaSynFlag As Integer

Const MAX_TREND_NUM As Integer = 100

Dim m_data(MAX_TREND_NUM)
Dim m_tempValue
Dim m_logoUID



Const MAX_IDLE_TIME_CALL As Integer = 3

Private Sub Class_Initialize()
    UpdateExpireDate
End Sub
'Give the formula to the related cell, The formula will be updated by Excell
Public Sub Initialize(LogoTrenduid As Integer, DirType As Integer, TrendNum As Integer, CallRange As Range)
    'MsgBox "Class_Initialize"
    Set m_WorkSheet = CallRange.Worksheet
     m_DirType = DirType
     m_logoUID = LogoTrenduid
     'now we use global m_DirNum. Maybe new requirement need to config the num for every formula
     m_TrendNum = TrendNum
     m_Row = CallRange.Row
     m_Column = CallRange.Column
    
     m_TrendChangeIndex = 1 'start from 1
     m_FormulaSynFlag = 0 ' not syn the formula, when first added and change the m_DirNum. It's 0
     DebugLog "[CllRangeContainer]m_Row" + CStr(m_Row) + "m_Column" + CStr(m_Column) + "Is added"
End Sub

Public Sub SetFormulaSynFLag()
   m_FormulaSynFlag = 0
End Sub



Public Function GetValueByIndex(index As Integer)
    GetValueByIndex = m_data(index)
End Function
Public Function ValidCallRange()
    Dim isValid As Boolean
    isValid = True
    Dim index As Integer
    Dim trend_cell As Range
    Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column)
     
    For index = 1 To m_DirNum Step 1
        Select Case m_DirType
            Case DIR_TR
                Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column + index)
            Case DIR_TL
              Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column - index)
            Case DIR_TU
                Set trend_cell = m_WorkSheet.Cells(m_Row - index, m_Column)
            Case DIR_TD
              Set trend_cell = m_WorkSheet.Cells(m_Row + index, m_Column)
        End Select
        tempFormula = "=GetTrendHisValue(" & CStr(m_logoUID) & "," & CStr(index) & ")"
        If InStr(trend_cell.Formula, tempFormula) <= 0 And InStr(trend_cell.Formula, "LOGOVAR") <= 0 Then
            trend_cell.Formula = tempFormula
        End If
    Next
    ValidCallRange = isValid
End Function
'just update the array
'when log out ,should be used
Public Sub RemoveTrendFormula()
    Dim index As Integer
    Dim trend_cell As Range
    For index = m_DirNum To 1 Step -1
        Select Case m_DirType
            Case DIR_TR
                Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column + index)
            Case DIR_TL
              Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column - index)
            Case DIR_TU
                Set trend_cell = m_WorkSheet.Cells(m_Row - index, m_Column)
            Case DIR_TD
              Set trend_cell = m_WorkSheet.Cells(m_Row + index, m_Column)
        End Select
        If InStr(trend_cell.Formula, "=GetTrendHisValue") > 0 Then
            trend_cell.ClearContents
        End If
    Next
End Sub
Public Sub UpdateHisdata()
       Dim temValue As Integer
        Dim ShowBegin As Single
        Dim ShowEnd As Single
     
     ' ShowBegin = Timer
           
    'insert the value of the last circle
        
  
        
        If Not IsError(m_WorkSheet.Cells(m_Row, m_Column)) Then
           m_data(0) = m_WorkSheet.Cells(m_Row, m_Column)
        End If
       'ShowEnd = Timer
            
       'Debug.Print "Fetch:"; ShowEnd - ShowBegin
 
    
       For tempValue = m_DirNum To 1 Step -1
            m_data(tempValue) = m_data(tempValue - 1)
        Next
        
  
End Sub
Public Sub AddFormula(TotalTrendNum As Integer, FinishedTrendNum As Integer)
    Dim tempRow As Integer
    Dim tempColumn As Integer
      Dim ShowBegin As Single
     Dim ShowEnd As Single
     Dim tempIndex As Integer
    tempRow = m_Row
    tempColumn = m_Column
    Dim CurentWorkSheet As Worksheet
    Dim objRange As Range
    Set CurentWorkSheet = m_WorkSheet
    If m_FormulaSynFlag <> 0 Then
        Exit Sub
    End If
     For m_TrendChangeIndex = 1 To m_DirNum Step 1
        prgramBarShow.Show
        Application.ScreenUpdating = False
        prgramBarShow.Repaint
        'sngBegin = Timer
        'Dim entry
           '按任务完成情况的百分比设置lblProgress的宽度，实现模拟进度效果
        prgramBarShow.lblprogress.Width = prgramBarShow.lblBack.Width * (FinishedTrendNum * m_DirNum + m_TrendChangeIndex) / (TotalTrendNum * m_DirNum)
            '格式化显示完成任务情况的百分比
        prgramBarShow.percert = Format((FinishedTrendNum * m_DirNum + m_TrendChangeIndex) / (TotalTrendNum * m_DirNum), "0%")
            '将进度自身重绘，实时显示
        prgramBarShow.Repaint
     tempFormula = "=GetTrendHisValue(" & CStr(m_logoUID) & "," & CStr(m_TrendChangeIndex) & ")"
    Select Case m_DirType
    'in CallTrendProcedure() the m_OldDirNum equal m_DirNum after one cycle.
     Case DIR_TR
     ' first time add the formula or  increase the dir num
        'If m_HisTrendFlag = 0 Then
            'ShowBegin = Timer
            Set objRange = CurentWorkSheet.Cells(tempRow, tempColumn + m_TrendChangeIndex)
            'ShowEnd = Timer
            'Debug.Print "Formulatime:"; ShowEnd - ShowBegin
    Case DIR_TL
            Set objRange = CurentWorkSheet.Cells(tempRow, tempColumn - m_TrendChangeIndex)
    Case DIR_TU
            Set objRange = CurentWorkSheet.Cells(tempRow - m_TrendChangeIndex, tempColumn)
    Case DIR_TD
            Set objRange = CurentWorkSheet.Cells(tempRow + m_TrendChangeIndex, tempColumn)
    End Select
     If InStr(objRange.Formula, "LOGOVAR") <= 0 Then
            objRange.Formula = tempFormula
    End If
  Next
  m_FormulaSynFlag = 1
End Sub
Public Sub DecreaceDirNum()
    Dim tempRow As Integer
    Dim tempColumn As Integer
    
      Dim ShowBegin As Single
     Dim ShowEnd As Single
    tempRow = m_Row
    tempColumn = m_Column
    Dim CurentWorkSheet As Worksheet
    Dim objRange As Range
    Set CurentWorkSheet = m_WorkSheet
        If m_OldDirNum > m_DirNum Then
    Select Case m_DirType
             Case DIR_TR
                    'ShowBegin = Timer
                    'CurentWorkSheet.Cells(tempRow, tempColumn + m_TrendChangeIndex).Formula = tempFormula
                    'ShowEnd = Timer
                    'Debug.Print "Formulatime:"; ShowEnd - ShowBegin
                   Set objRange = CurentWorkSheet.Range(CurentWorkSheet.Cells(tempRow, tempColumn + m_DirNum + 1), CurentWorkSheet.Cells(tempRow, tempColumn + m_OldDirNum))
            Case DIR_TL
                    Set objRange = CurentWorkSheet.Range(CurentWorkSheet.Cells(tempRow, tempColumn - m_OldDirNum), CurentWorkSheet.Cells(tempRow, tempColumn - m_DirNum - 1))
            Case DIR_TU
                    Set objRange = CurentWorkSheet.Range(CurentWorkSheet.Cells(tempRow - m_OldDirNum, tempColumn), CurentWorkSheet.Cells(tempRow - m_DirNum, tempColumn - 1))
            Case DIR_TD
                    Set objRange = CurentWorkSheet.Range(CurentWorkSheet.Cells(tempRow + m_DirNum + 1, tempColumn), CurentWorkSheet.Cells(tempRow + m_OldDirNum, tempColumn))
    End Select
            objRange.ClearContents
        End If
End Sub

Public Sub TickIdleTime()
    m_executionTime = m_executionTime + 1
End Sub

Public Function GetDirType()
    GetDirType = m_DirType
End Function

Public Function IsExpired()
    Dim trend_cell As Range
    Set trend_cell = m_WorkSheet.Cells(m_Row, m_Column)
    IsExpired = InStr(trend_cell.Formula, "LOGOVAR") <= 0 Or m_executionTime > MAX_IDLE_TIME_CALL
End Function

' ==================== Private sections ===========================

Public Sub UpdateExpireDate()
    m_executionTime = 0 ' reset execution time when it is referred

    'm_expireDate = Now + TimeValue("00:00:10") ' expires when it is not used in 10 seconds
End Sub


