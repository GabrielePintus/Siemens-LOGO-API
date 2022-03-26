VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariablePackerV81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Configurations according to connection devices
Dim DEVICE_TYPE As Integer
Dim MAX_REQUEST_SIZE As Integer
Dim MAX_RESPONSE_SIZE As Integer
Dim MAX_SECTION_SIZE As Integer
Dim RESPONSE_ASCII As Boolean

Dim request As String
Dim requestSize As Integer
Dim responseSize As Integer
Dim requestParts As DArray

Dim Range As Integer
Dim rangeByteStart As Integer
Dim rangeByteEnd As Integer

Private Sub Class_Initialize()
    DEVICE_TYPE = GetConnectionType()
    Set requestParts = New DArray
    
    ' Update configurations accordingly
    Select Case DEVICE_TYPE
    Case CONNECTION_0BA8
        MAX_REQUEST_SIZE = 1024
        MAX_RESPONSE_SIZE = 1015
        MAX_SECTION_SIZE = 4
        RESPONSE_ASCII = False
    Case Else
        MAX_REQUEST_SIZE = 1024
        MAX_RESPONSE_SIZE = 1015
        MAX_SECTION_SIZE = 128
        RESPONSE_ASCII = True
    End Select
End Sub

Private Function GetRequestString(Range, byteStart, byteEnd)
    Select Case DEVICE_TYPE
    Case CONNECTION_0BA8
        Dim byteCount
        byteCount = byteEnd - byteStart + 1
        
        If byteCount <= 1 Then
            GetRequestString = "v," + CStr(Range) + "," + CStr(byteStart * 8) + ",2" ' BYTE
        Else
            If byteCount <= 2 Then
                GetRequestString = "," + CStr(Range) + "," + CStr(byteStart * 8) + ",4" ' WORD
            Else
                GetRequestString = "," + CStr(Range) + "," + CStr(byteStart * 8) + ",6" ' DWORD
            End If
        End If
    Case Else
        GetRequestString = "," + CStr(Range) + ",0," + CStr(byteStart * 8) + ",2," + CStr(byteEnd - byteStart + 1)
    End Select
End Function

Private Function GetRequestSize(Range, byteStart, byteEnd)
    Select Case DEVICE_TYPE
    Case CONNECTION_0BA8
        'v,range,addr,type (without ; delimiter)
        GetRequestSize = 5 + GetNumberDigits(Range) + GetNumberDigits(byteStart * 8)
    Case Else
        ',range,0,addr,2,length (without ; delimiter)
        GetRequestSize = 7 + GetNumberDigits(Range) + GetNumberDigits(byteStart * 8) + GetNumberDigits(byteEnd - byteStart + 1)
    End Select
End Function

Private Function GetResponseSize(Range, byteStart, byteEnd)
    Select Case DEVICE_TYPE
    Case CONNECTION_0BA8
        '<r i='' v='' />
        If byteCount <= 1 Then
            GetResponseSize = 15 + 3 ' BYTE max 255
        Else
            If byteCount <= 2 Then
                GetResponseSize = 15 + 5 ' WORD max 65535
            Else
                GetResponseSize = 15 + 10 ' DWORD max 4294967295
            End If
        End If
    Case Else
        '<r i='' e='0' v= '00'/>
        GetResponseSize = 20 + (byteEnd - byteStart + 1) * 2
    End Select
End Function

Private Function GetNumberDigits(num)
    If num < 10 Then
        GetNumberDigits = 1
        Exit Function
    End If
    
    If num < 100 Then
        GetNumberDigits = 2
        Exit Function
    End If
    
    If num < 1000 Then
        GetNumberDigits = 3
        Exit Function
    End If
    
    If num < 10000 Then
        GetNumberDigits = 4
        Exit Function
    End If
    
    GetNumberDigits = 8
End Function

Private Sub Flush()
    If Range <> 0 Then
        Dim entryRequestStr
        entryRequestStr = GetRequestString(Range, rangeByteStart, rangeByteEnd)
        
        If request <> "" Then
            request = request + ";"
            requestSize = requestSize + 1
        End If
        
        request = request + entryRequestStr
        requestSize = requestSize + Len(entryRequestStr)
        responseSize = responseSize + GetResponseSize(Range, rangeByteStart, rangeByteEnd)
        
        ' Record the item
        Dim item As PackItem
        
        Set item = New PackItem
        item.Range = Range
        item.byteStart = rangeByteStart
        item.byteEnd = rangeByteEnd
        
        requestParts.Add item
        
        ' clear current item
        Range = 0
        rangeByteStart = 0
        rangeByteEnd = 0
    End If
End Sub

Public Function GetRequestStr()
    ' always flush before returning the request
    Flush
    
    GetRequestStr = request
End Function

Public Function CanAcceptTheRange(r, S, e)
    ' Don't accept a big range
    If e - S + 1 > MAX_SECTION_SIZE Then
        CanAcceptTheRange = False
        Exit Function
    End If

    ' Don't accept a big request
    If requestSize + GetRequestSize(r, S, e) + 1 > MAX_REQUEST_SIZE Then ' Plus probably 1 delimiter
        CanAcceptTheRange = False
        Exit Function
    End If
    
    ' Don't accept a big response
    If responseSize + GetResponseSize(r, S, e) > MAX_RESPONSE_SIZE Then
        CanAcceptTheRange = False
        Exit Function
    End If
    
    CanAcceptTheRange = True
End Function

Public Function AddVariable(entry)
    Dim entryRange As Integer
    Dim entryByteStart As Integer
    Dim entryByteEnd As Integer
        
    entryRange = entry.GetRange()
    entryByteStart = entry.GetAddress() \ 8
    entryByteEnd = (entry.GetAddress() + entry.GetBitsSize() - 1) \ 8 ' the last valid one! (not the next invalid one!)

    ' Handle original range (if it exists). Merge and quit
    If Range = entryRange Then
        ' This is possible when it has a short length
        If entryByteEnd < rangeByteEnd Then
            AddVariable = True
            Exit Function
        End If
    
        ' If there is a small gap, merge it
        If (entryByteStart - rangeByteEnd - 1) <= MAX_SIZE_GAP Then
            If CanAcceptTheRange(Range, rangeByteStart, entryByteEnd) Then
                rangeByteEnd = entryByteEnd
                AddVariable = True
                Exit Function
            End If
        End If
    End If
    
    ' If it cannot be merged, always try to flush original range at first
    Flush
    
    ' Then tries to append new entry
    If CanAcceptTheRange(entryRange, entryByteStart, entryByteEnd) Then
        Range = entryRange
        rangeByteStart = entryByteStart
        rangeByteEnd = entryByteEnd
        AddVariable = True
        Exit Function
    End If

    AddVariable = False
End Function

Public Function GetData(xmlRoot)
    Dim data As VariableData
    
    Set data = New VariableData

    If requestParts.Count = xmlRoot.Length Then
        Dim i
        For i = 0 To xmlRoot.Length - 1 Step 1
            Dim item As PackItem
            Set item = requestParts.GetItem(i)
            
            data.FillData item.Range, item.byteStart, item.byteEnd, xmlRoot.item(i).getAttribute("v"), RESPONSE_ASCII
        Next
    Else
        'MsgBox "invalid response"
    End If
    
    Set GetData = data
End Function
