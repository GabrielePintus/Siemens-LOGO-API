VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariableEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_reqId
Dim m_val
'Dim m_expireDate
Dim m_reqStr
Dim m_executionTime As Integer
Public m_Range As Integer
Public m_addr As Integer
Public m_size As Integer

Const MAX_IDLE_TIME As Integer = 3

Private Sub Class_Initialize()
    UpdateExpireDate
End Sub

Public Sub Initialize(Range As Integer, Addr As Integer, size As Integer)
    m_Range = Range
    m_addr = Addr
    m_size = size
    m_reqStr = CStr(Range) + ",0," + CStr(Addr) + "," + CStr(size) + ",1" 'the additional 0 and 1 is the subrange for V8.1
End Sub

Public Function GetReqStr()
    GetReqStr = m_reqStr
End Function

Public Function GetBitsSize()
    Select Case m_size
    Case VAR_SIZE_BIT
        GetBitsSize = 1
    Case VAR_SIZE_BYTE
        GetBitsSize = 8
    Case VAR_SIZE_WORD
        GetBitsSize = 16
    Case VAR_SIZE_DWORD
        GetBitsSize = 32
    Case Else
        GetBitsSize = 0
    End Select
End Function

Public Function GetAddress()
    GetAddress = m_addr
End Function

Public Function GetRange()
    GetRange = m_Range
End Function

Public Sub SetReqId(reqId)
    m_reqId = reqId
End Sub

Public Function GetReqId()
    GetReqId = m_reqId
End Function

Public Function GetValue()
    UpdateExpireDate

    GetValue = m_val
    'GetValue = GetVarValue(Me)
End Function

Public Sub UpdateValue(val)
    m_val = val
End Sub

Public Sub TickIdleTime()
    m_executionTime = m_executionTime + 1
End Sub

Public Function IsExpired()
    IsExpired = m_executionTime > MAX_IDLE_TIME
End Function

' ==================== Private sections ===========================

Private Sub UpdateExpireDate()
    m_executionTime = 0 ' reset execution time when it is referred

    'm_expireDate = Now + TimeValue("00:00:10") ' expires when it is not used in 10 seconds
End Sub

