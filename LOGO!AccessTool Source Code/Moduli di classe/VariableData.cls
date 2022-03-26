VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariableData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_ranges

Private Sub Class_Initialize()
    Set m_ranges = CreateObject("Scripting.Dictionary")
End Sub

Private Function ParseHexStr(STR)
    Dim val
    val = CLng("&H" + STR)
    If val < 0 Then
        val = 4294967296# + val
    End If
    ParseHexStr = val
End Function

Public Sub FillData(Range, byteStart, byteEnd, ascData, isAscData)
    If ascData <> "" Then
        Dim rangeObj
        If Not m_ranges.exists(Range) Then
            Set rangeObj = CreateObject("Scripting.Dictionary")
            m_ranges.Add Range, rangeObj
        Else
            Set rangeObj = m_ranges.item(Range)
        End If
    
        Dim dataLen
        dataLen = Len(ascData) \ 2
        
        Dim i
        For i = 0 To dataLen - 1
            rangeObj.item((byteStart + i)) = ParseHexStr(Mid(ascData, i * 2 + 1, 2))
        Next
    End If
End Sub

Public Function GetValue(entry)
    Dim Range
    Dim bitAddr
    Dim size

    Range = entry.GetRange()
    bitAddr = entry.GetAddress()
    size = entry.GetBitsSize()
    
    Dim rangeObj
    
    If Not m_ranges.exists(Range) Then
        GetValue = Empty
        Exit Function
    End If
    
    Set rangeObj = m_ranges.item(Range)
    
    Select Case size
    Case 1
        GetValue = GetBit(rangeObj, bitAddr)
    Case 8
        GetValue = GetByte(rangeObj, bitAddr)
    Case 16
        GetValue = GetWord(rangeObj, bitAddr)
    Case 32
        GetValue = GetDWord(rangeObj, bitAddr)
    Case Else
        GetValue = Empty
    End Select
End Function

Private Function GetBit(rangeObj, bitAddr)
    Dim byteAddr
    byteAddr = bitAddr \ 8
    
    If Not rangeObj.exists(byteAddr) Then
        GetBit = Empty
        Exit Function
    End If
        
    GetBit = TestBit(rangeObj.item(byteAddr), bitAddr Mod 8)
End Function

Private Function GetByte(rangeObj, bitAddr)
    If bitAddr Mod 8 <> 0 Then
        GetByte = Empty
        Exit Function
    End If
    
    Dim byteAddr
    byteAddr = bitAddr \ 8
    
    If Not rangeObj.exists(byteAddr) Then
        GetByte = Empty
        Exit Function
    End If
        
    GetByte = rangeObj.item(byteAddr)
End Function

Private Function GetWord(rangeObj, bitAddr)
    If bitAddr Mod 8 <> 0 Then
        GetWord = Empty
        Exit Function
    End If
    
    Dim byteAddr
    byteAddr = bitAddr \ 8
    
    If (Not rangeObj.exists(byteAddr)) Or (Not rangeObj.exists(byteAddr + 1)) Then
        GetWord = Empty
        Exit Function
    End If
    
    GetWord = rangeObj.item(byteAddr) * 256 + rangeObj.item(byteAddr + 1)
End Function

Private Function GetDWord(rangeObj, bitAddr)
    If bitAddr Mod 8 <> 0 Then
        GetDWord = Empty
        Exit Function
    End If
    
    Dim byteAddr
    byteAddr = bitAddr \ 8
    
    If (Not rangeObj.exists(byteAddr)) Or (Not rangeObj.exists(byteAddr + 1)) Or (Not rangeObj.exists(byteAddr + 2)) Or (Not rangeObj.exists(byteAddr + 3)) Then
        GetDWord = Empty
        Exit Function
    End If
    
    GetDWord = rangeObj.item(byteAddr) * 16777216 + rangeObj.item(byteAddr + 1) * 65536 + rangeObj.item(byteAddr + 2) * 256 + rangeObj.item(byteAddr + 3)
End Function

