Attribute VB_Name = "Common"
Option Private Module
Dim heartbeating As Boolean
Dim heartbeatHandlers

'Public Const VAR_SIZE_BIT = 1
'Public Const VAR_SIZE_BYTE = 2
'Public Const VAR_SIZE_WORD = 4
'Public Const VAR_SIZE_DWORD = 6
 Public Const VAR_UNSIGNED = 0
 Public Const VAR_SIGNED = 1
 Public Const VAR_HEX = 2
 Public Const VAR_BINARY = 3
 
 
 
 Public Const DIR_TR = 1
 Public Const DIR_TL = 2
 Public Const DIR_TU = 3
 Public Const DIR_TD = 4
 

 

Public Sub Test()



    'MsgBox TypeName(Selection)
    'Selection.Select
    MsgBox DataConvert(Empty, VAR_UNSIGNED, TestTe())
End Sub

Public Function TestTe()
     
    TestTe = 1
End Function


'This will be used by different LOGOVAR(L), The defaulst is unsigned
Public Function DataConvert(data, datatype As Integer, datasize As Integer)


    
    If datatype = VAR_HEX Then
        DataConvert = D2H(data)
         Exit Function
    End If
  'problem?
    If datatype = VAR_BINARY Then
         DataConvert = D2B(data)
         Exit Function
    End If
    '1bit unsinged ,no need to change
    If (datasize = VAR_SIZE_BIT) Or (datatype = VAR_UNSIGNED) Then
        DataConvert = data
        Exit Function
    End If
    
    If datatype = VAR_SIGNED Then
        Select Case datasize
        Case 8 'VAR_SIZE_BYTE
            If (CByte(data) > 127) Then
            DataConvert = CByte(data) - 255 - 1
            Else
            DataConvert = data
            End If
        Case 16 'VAR_SIZE_WORD
            If (data > 32767) Then
            DataConvert = data - 65535 - 1
            Else
            DataConvert = data
            End If
           
        Case 32 'VAR_SIZE_DWORD
            If (data > 2147483647#) Then
            DataConvert = data - 4294967295# - 1
            Else
            DataConvert = data
            End If
        End Select
    
    End If
 
  
End Function




Public Function D2H(Dec) As String
     Dim a As String
     D2H = ""
     Dim msb As Integer
     Dim lowlong As Long
     Dim i
     msb = 0
     
     If Dec > 2147483647 Then
            msb = 1
            lowlong = Dec - 2147483648#
           
     Else
        lowlong = Dec
     End If
      For i = 0 To 7
           If (msb = 1) And (i = 7) Then
                a = CStr((lowlong Mod 16) + 8)
           Else
                a = CStr(lowlong Mod 16)
           End If
           
           Select Case a
             Case "10": a = "A"
             Case "11": a = "B"
             Case "12": a = "C"
             Case "13": a = "D"
             Case "14": a = "E"
             Case "15": a = "F"
            End Select
            D2H = a & D2H
            
            lowlong = lowlong \ 16
            If (lowlong = 0) And (msb = 0) Then
                 Exit For
            End If
      Next
   
End Function

Public Function D2B(Dec) As String
     D2B = ""
     Dim msb As Integer
     Dim lowlong As Long
     Dim i
     msb = 0
     If Dec > 2147483647 Then
        msb = 1
        lowlong = Dec - 2147483648#
     Else
        lowlong = Dec
     End If
    For i = 0 To 31
      If (msb = 1) And (i = 31) Then
           D2B = 1 & D2B

      Else
           D2B = (lowlong Mod 2) & D2B
           lowlong = lowlong \ 2
      End If
      
      If (lowlong = 0) And (msb = 0) Then
            Exit For
      End If
    Next
End Function

Public Sub Test34()
    MsgBox Application.Caller
End Sub

Public Function GetUnusedName(prefix, Workbook As Workbook)
    Dim shapeObj

    Set shs = Workbook.Sheets ' Application.Sheets may access xlam's sheets
    
    While True
        Randomize
        Dim X
        X = Int((1000 * Rnd) + 1)
        Dim groupName
        groupName = prefix + CStr(X)
        
        Dim found
        found = False
        
        Dim i
        
        For i = 1 To shs.Count
            Dim j
            For j = 1 To shs(i).Shapes.Count
                Set shapeObj = shs(i).Shapes.item(j)
                
                If shapeObj.name = groupName Then
                    found = True
                    Exit For ' stop trying, just start another round
                End If
            Next
            
            If found Then
                Exit For ' stop trying, just start another round
            End If
        Next
        
        If Not found Then
            GetUnusedName = groupName
            Exit Function
        End If
    Wend
End Function

Public Sub TestTimer()
    MsgBox heartbeating
     
    MsgBox TypeName(heartbeatHandlers)
End Sub

Public Function GetHeartBeatHandlers()
    If TypeName(heartbeatHandlers) = "Empty" Then
        Set heartbeatHandlers = CreateObject("Scripting.Dictionary")
        
        ' Call this interface to initialize environment
        InitializeEnvironment
    End If
    
    If Not heartbeating And heartbeatHandlers.Count > 0 Then
        heartbeating = True
        
        Application.OnTime Now + TimeValue("00:00:01"), "OnHeartBeatTimer" ' Try to start heartbeating
    End If
    
    Set GetHeartBeatHandlers = heartbeatHandlers
End Function

Public Sub RegisterHeartBeatTimer(handler)
    Dim handlers
    Set handlers = GetHeartBeatHandlers
    
    If Not handlers.exists(handler) Then
        handlers.Add handler, 0
        'm_StopFlagChange = 1 'used to help make react more quick.
    End If
    
    ' Try start heartbeating after registered new timer
    If Not heartbeating And heartbeatHandlers.Count > 0 Then
       
        heartbeating = True
        
        Application.OnTime Now + TimeValue("00:00:01"), "OnHeartBeatTimer" ' Try to start heartbeating
    End If
End Sub

Public Sub UnRegisterHeartBeatTimer(handler)
    Dim handlers
    Set handlers = GetHeartBeatHandlers
    
    If handlers.exists(handler) Then
        handlers.Remove handler
    End If
End Sub

' The timer may not be stopped when engine stops. However, the variable environment does
' This handler shall prevent from multiple timing
Public Sub OnHeartBeatTimer()
   Dim handlers
    Set handlers = GetHeartBeatHandlers
    
    ' Implicit periodical Task -- Validate environment at each heartbeat incase the document is closed
    ' ValidateEnvironment
    
    If handlers.Count > 0 Then
        heartbeating = True
    
        Application.OnTime Now + TimeValue("00:00:01"), "OnHeartBeatTimer" ' register for another heart beat
        
        Dim keys
        keys = handlers.keys
        
        ' Then let each handlers operate
        Dim i
        For i = 0 To handlers.Count - 1
            Application.Run keys(i)
        Next
     
    Else
        heartbeating = False ' stop heart beating if no handlers defined
    End If
End Sub

' The general error handler (After connected). It can be used for all AJAX requests after login
Public Sub CommonOnConnectionError(arg0, status)
    ' Do nothing if it is not successfully connected. This could avoid batch error promption
    Dim currentStatus
    currentStatus = GetProperty(PROPERTY_ID_CONNECTION)
    
    ' Only handle error in connected case. Error in other state will be simply ignored
    If currentStatus = STATE_CONNECTED Then
        If status = 403 Then
             MsgBox "CommonOnConnectionError 403 error"
            RecoverConnection
        Else
            LogOut
        
            MsgBox STR(MSG_CONN_BROKEN) ' Notify user about connection broken (timeout or http error)"
        End If
    End If
End Sub

Public Sub CommonOnConnectionRecoverFail(arg0, status)
    ' Do nothing if it is not successfully connected. This could avoid batch error promption
    Dim currentStatus
    currentStatus = GetProperty(PROPERTY_ID_CONNECTION)
    
    If currentStatus = STATE_RECOVERING Then
        MsgBox STR(MSG_CONN_BROKEN) ' Notify user about connection broken (timeout or http error)"
    End If
End Sub

