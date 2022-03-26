Attribute VB_Name = "LOGOLOG"
'Public Const VAR_SIZE_BIT = 1
'Public Const VAR_SIZE_BYTE = 2
'Public Const VAR_SIZE_WORD = 4
'Public Const VAR_SIZE_DWORD = 6
' Public Const VAR_UNSIGNED = 0
 'Public Const VAR_SIGNED = 1
 'Public Const VAR_HEX = 2
 'Public Const VAR_BINARY = 3
 
 ' Public Const DIR_TR = 1
 'Public Const DIR_TL = 2
' Public Const DIR_TU = 3
 'Public Const DIR_TD = 4
 
  
Public myc
 


' new added 14*2 function for trend
Public Function LOGOVAR(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVAR = logovarlog(id, VAR_UNSIGNED, 0, DIR_TR)
     Case "TD"
    LOGOVAR = logovarlog(id, VAR_UNSIGNED, 0, DIR_TD)
     Case "NODIR"
    LOGOVAR = logovarlog(id, VAR_UNSIGNED, 0)
     Case Else
     LOGOVAR = Nothing
     End Select
End Function


Public Function LOGOVARU(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARU = logovarlog(id, VAR_UNSIGNED, 0, DIR_TR)
     Case "TD"
    LOGOVARU = logovarlog(id, VAR_UNSIGNED, 0, DIR_TD)
     Case "NODIR"
    LOGOVARU = logovarlog(id, VAR_UNSIGNED, 0)
     Case Else
    LOGOVARU = Nothing
     End Select
End Function

Public Function LOGOVARS(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARS = logovarlog(id, VAR_SIGNED, 0, DIR_TR)
     Case "TD"
    LOGOVARS = logovarlog(id, VAR_SIGNED, 0, DIR_TD)
     Case "NODIR"
    LOGOVARS = logovarlog(id, VAR_SIGNED, 0)
     Case Else
    LOGOVARS = Nothing
     End Select
End Function

Public Function LOGOVARH(id As String, Optional tr As String = "NODIR")
     Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARH = logovarlog(id, VAR_HEX, 0, DIR_TR)
     Case "TD"
    LOGOVARH = logovarlog(id, VAR_HEX, 0, DIR_TD)
     Case "NODIR"
    LOGOVARH = logovarlog(id, VAR_HEX, 0)
     Case Else
    LOGOVARH = Nothing
     End Select
End Function

Public Function LOGOVARB(id As String, Optional tr As String = "NODIR")
 Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARB = logovarlog(id, VAR_BINARY, 0, DIR_TR)
     Case "TD"
    LOGOVARB = logovarlog(id, VAR_BINARY, 0, DIR_TD)
     Case "NODIR"
    LOGOVARB = logovarlog(id, VAR_BINARY, 0)
     Case Else
    LOGOVARB = Nothing
     End Select
End Function

Public Function LOGOVARL(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARL = logovarlog(id, VAR_UNSIGNED, 1, DIR_TR)
     Case "TD"
    LOGOVARL = logovarlog(id, VAR_UNSIGNED, 1, DIR_TD)
     Case "NODIR"
    LOGOVARL = logovarlog(id, VAR_UNSIGNED, 1)
     Case Else
    LOGOVARL = Nothing
     End Select
End Function

Public Function LOGOVARUL(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARUL = logovarlog(id, VAR_UNSIGNED, 1, DIR_TR)
     Case "TD"
    LOGOVARUL = logovarlog(id, VAR_UNSIGNED, 1, DIR_TD)
     Case "NODIR"
    LOGOVARUL = logovarlog(id, VAR_UNSIGNED, 1)
     Case Else
    LOGOVARUL = Nothing
     End Select
End Function

Public Function LOGOVARSL(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARSL = logovarlog(id, VAR_SIGNED, 1, DIR_TR)
     Case "TD"
    LOGOVARSL = logovarlog(id, VAR_SIGNED, 1, DIR_TD)
     Case "NODIR"
    LOGOVARSL = logovarlog(id, VAR_SIGNED, 1)
     Case Else
    LOGOVARSL = Nothing
     End Select
End Function

Public Function LOGOVARBL(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARBL = logovarlog(id, VAR_BINARY, 1, DIR_TR)
     Case "TD"
    LOGOVARBL = logovarlog(id, VAR_BINARY, 1, DIR_TD)
     Case "NODIR"
    LOGOVARBL = logovarlog(id, VAR_BINARY, 1)
     Case Else
    LOGOVARBL = Nothing
     End Select
End Function

Public Function LOGOVARHL(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARHL = logovarlog(id, VAR_HEX, 1, DIR_TR)
     Case "TD"
    LOGOVARHL = logovarlog(id, VAR_HEX, 1, DIR_TD)
     Case "NODIR"
    LOGOVARHL = logovarlog(id, VAR_HEX, 1)
     Case Else
    LOGOVARHL = Nothing
     End Select
End Function


Public Function LOGOVARLU(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARLU = logovarlog(id, VAR_UNSIGNED, 1, DIR_TR)
     Case "TD"
    LOGOVARLU = logovarlog(id, VAR_UNSIGNED, 1, DIR_TD)
     Case "NODIR"
    LOGOVARLU = logovarlog(id, VAR_UNSIGNED, 1)
     Case Else
    LOGOVARLU = Nothing
     End Select
End Function

Public Function LOGOVARLS(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARLS = logovarlog(id, VAR_SIGNED, 1, DIR_TR)
     Case "TD"
    LOGOVARLS = logovarlog(id, VAR_SIGNED, 1, DIR_TD)
     Case "NODIR"
    LOGOVARLS = logovarlog(id, VAR_SIGNED, 1)
     Case Else
    LOGOVARLS = Nothing
     End Select
End Function

Public Function LOGOVARLB(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARLB = logovarlog(id, VAR_BINARY, 1, DIR_TR)
     Case "TD"
    LOGOVARLB = logovarlog(id, VAR_BINARY, 1, DIR_TD)
     Case "NODIR"
    LOGOVARLB = logovarlog(id, VAR_BINARY, 1)
     Case Else
    LOGOVARLB = Nothing
     End Select
End Function

Public Function LOGOVARLH(id As String, Optional tr As String = "NODIR")
    Dim m_trend
    m_trend = UCase(tr)
    Select Case m_trend
      Case "TR"
    LOGOVARLH = logovarlog(id, VAR_HEX, 1, DIR_TR)
     Case "TD"
    LOGOVARLH = logovarlog(id, VAR_HEX, 1, DIR_TD)
     Case "NODIR"
    LOGOVARLH = logovarlog(id, VAR_HEX, 1)
     Case Else
    LOGOVARLH = Nothing
     End Select
End Function

' only called by the log related UI function
Private Function logovarlog(id As String, datatype As Integer, logneed As Integer, Optional DirType As Integer = 0)
    On Error GoTo Err
    'MsgBox "test"
  
  
     
    Dim tempCallRange As Range
   
    
    If DirType <> 0 And TypeName(Application.Caller) = "Range" Then
        Set tempCallRange = Application.Caller
        AddCallRange DirType, m_Interval, tempCallRange
     End If
    

    
  
  
  
    'for datatype, 0 1 2 refer to diff num type
    Application.Volatile
    Dim val
    Dim typevalue
    Dim formattedId
    Dim tempValue
    formattedId = UCase(id)
    Dim entry As Object
   
   'AddCallRange CallRange, 1, 5
    
    
    If logneed = 1 Then
    
      
      Dim Column
      Column = 0 ' Make it 0 by default
      
      ' Try to parse column from formattedId
      Dim pos
      pos = InStr(formattedId, "@")
      
      If pos > 0 Then
          Dim columnStr
          columnStr = Trim(Mid(formattedId, pos + 1))
          If columnStr <> "" Then
              Column = CDbl(columnStr)
          End If
          
          formattedId = Left(formattedId, pos - 1)
      End If
    
        
         

      ' Get workbook container
      Dim workbookContainer As Object
      If TypeName(Application.Caller) = "Range" Then
          Set workbookContainer = GetWorkBookContainer(Application.Caller.Worksheet.Parent)
      End If
      
      ' try to get value and entry
    
      
      Dim dataentry As Object
      
      Select Case formattedId
      Case "STATUS"
          typevalue = GetStatusString()
          If Not workbookContainer Is Nothing Then
              workbookContainer.AddVARL formattedId, Column, typeentry ' Record it in DataLog collector
          End If
          
      Case "TIME"
            typevalue = GetCurTime()
      Case Else
          Set entry = GetVariableEntry(formattedId, False) ' Don't add invalid one
          
          If Not entry Is Nothing Then
             val = entry.GetValue()
             typevalue = DataConvert(val, datatype, entry.GetBitsSize())
             
             
             'If getStopFlag <> 1 Then
                Set dataentry = New TypeVariableEntry
                dataentry.Initialize entry.GetRange(), entry.GetAddress(), entry.GetBitsSize(), datatype
                dataentry.UpdateValue typevalue
                If Not workbookContainer Is Nothing Then
                     If Not IsEmpty(typevalue) Then
                        workbookContainer.AddVARL formattedId, Column, dataentry ' Record it in DataLog collector
                    End If
                 End If
              'End If
          Else
              ' val = "Invalid logovarlog" ' invalid VARL shall not be recorded
              typevalue = Empty ' Make it the same as LOGOVAR
          End If
      End Select
    Else ' this is for no datalog
        
    
        Select Case formattedId
        Case "STATUS"
            typevalue = GetStatusString()
        Case "TIME"
            typevalue = GetCurTime()
        Case Else

            Set entry = GetVariableEntry(id, True) ' When it is really asked to get value, add it even invalid
            
            If Not entry Is Nothing Then
                val = entry.GetValue()
                'typevalue = Test341(val, datatype, entry.GetBitsSize())
                'MsgBox entry.GetSize()
                typevalue = DataConvert(val, datatype, entry.GetBitsSize())
                'Test34
            Else
                typevalue = Empty ' this will make it invalid value. use it as default
            End If
        End Select
        
    End If ' this is for no datalog
      
    'Set myc = New css
    'Set myc.sht = ActiveSheet
    
    If getStopFlag = 1 Then ' Stop state
        Set logovarlog = Nothing
    ElseIf TypeName(typevalue) = "Empty" Then
        Set logovarlog = Nothing ' this will make it invalid value (#VALUE)
    Else
        logovarlog = typevalue
    End If
     
    
   'Set myc = New css
    
   'Set myc.sht = ActiveSheet
 
    
    Exit Function
    
Err:
        logovarlog = "Invalid logovarlog" ' return invalid logovarlog to indicate error

End Function


