VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariablePackerV8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DEVICE_TYPE As Integer
Dim MAX_REQUEST_SIZE As Integer
Dim MAX_RESPONSE_SIZE As Integer
Dim MAX_SECTION_SIZE As Integer
Dim RESPONSE_ASCII As Boolean

Dim request As String
Dim requestSize As Integer
Dim responseSize As Integer

Private Sub Class_Initialize()
    DEVICE_TYPE = GetConnectionType()
    
    ' Update configurations accordingly
    MAX_REQUEST_SIZE = 1400 ' 2400 - 1000(head)
    MAX_RESPONSE_SIZE = 2181 ' (2400 - 210(head) - 9(response xml))
    MAX_SECTION_SIZE = 4
    RESPONSE_ASCII = False
End Sub

Public Function AddVariable(entry)
    Dim entryReqId
    Dim entryReqStr
    Dim entryResponseSize

    entryReqId = entry.GetReqId()
    entryReqStr = entryReqId + "," + CStr(entry.GetRange()) + "," + CStr(entry.GetAddress())
    entryResponseSize = 15 + Len(entryReqId) '<r i='' v='' />
    
    Select Case entry.GetBitsSize()
    Case 1
        entryReqStr = entryReqStr + ",1" ' BIT
        entryResponseSize = entryResponseSize + 1  ' Bit max 1
    Case 8
        entryReqStr = entryReqStr + ",2" ' BYTE
        entryResponseSize = entryResponseSize + 3 ' BYTE max 255
    Case 16
        entryReqStr = entryReqStr + ",3" ' WORD
        entryResponseSize = entryResponseSize + 5 ' WORD max 65535
    Case 32
        entryReqStr = entryReqStr + ",4" ' DWORD
        entryResponseSize = entryResponseSize + 10 ' DWORD max 4294967295
    Case Else
        ' Not necessary to handle it. Simply ignore it
        AddVariable = True
        Exit Function
    End Select
    

    ' Don't accept a big request
    If requestSize + Len(entryReqStr) + 1 > MAX_REQUEST_SIZE Then ' Plus probably 1 delimiter
        AddVariable = False
        Exit Function
    End If
    
    ' Don't accept a big response
    If responseSize + entryResponseSize > MAX_RESPONSE_SIZE Then
        AddVariable = False
        Exit Function
    End If
    
    ' It is OK to append the request
    If request <> "" Then
        request = request + ";"
        requestSize = requestSize + 1
    End If
    
    request = request + entryReqStr
    requestSize = requestSize + Len(entryReqStr)
    responseSize = responseSize + entryResponseSize

    AddVariable = True
End Function

Public Function GetRequestStr()
    GetRequestStr = request
End Function

Public Function GetData(xmlRoot)
    Dim data
    
    Set data = New VariableDataV8
    
    Dim i
    For i = 0 To xmlRoot.Length - 1 Step 1
        data.FillData xmlRoot.item(i).getAttribute("i"), xmlRoot.item(i).getAttribute("v"), RESPONSE_ASCII
    Next
    
    Set GetData = data
End Function

