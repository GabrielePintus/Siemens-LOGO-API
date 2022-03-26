VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariableFactoryV81"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Create2DVariableEntry(Range As String, Addr As Integer, subAddr As Integer)
    Set Create2DVariableEntry = Nothing ' Return Nothing by default to indicate error
    
    Select Case Range
    Case "SR"
        If ValidateNumber(Addr, 1, 4) And ValidateNumber(subAddr, 1, 8) Then
            Set Create2DVariableEntry = NewVariableEntry(14, (Addr - 1) * 8 + (subAddr - 1), VAR_SIZE_BIT)
        End If
    Case "VB"
        If ValidateNumber(Addr, 0, 850) And ValidateNumber(subAddr, 0, 7) Then
            Set Create2DVariableEntry = NewVariableEntry(132, Addr * 8 + subAddr, VAR_SIZE_BIT)
        End If
    Case Else
    End Select
End Function

Public Function CreateVariableEntry(Range As String, Addr As Integer)
    Set CreateVariableEntry = Nothing ' Return Nothing by default to indicate error
    
    Select Case Range
    Case "VB"
        If ValidateNumber(Addr, 0, 850) Then
            Set CreateVariableEntry = NewVariableEntry(132, Addr * 8, VAR_SIZE_BYTE) 'VM address starts from 0
        End If
    Case "VW"
        If ValidateNumber(Addr, 0, 849) Then
            Set CreateVariableEntry = NewVariableEntry(132, Addr * 8, VAR_SIZE_WORD) 'VM address starts from 0
        End If
    Case "VD"
        If ValidateNumber(Addr, 0, 847) Then
            Set CreateVariableEntry = NewVariableEntry(132, Addr * 8, VAR_SIZE_DWORD) 'VM address starts from 0
        End If
    Case "I"
        If ValidateNumber(Addr, 1, 24) Then
            Set CreateVariableEntry = NewVariableEntry(129, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "NI"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(16, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "Q"
        If ValidateNumber(Addr, 1, 20) Then
            Set CreateVariableEntry = NewVariableEntry(130, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "NQ"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(17, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "M"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(131, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "AI"
        If ValidateNumber(Addr, 1, 8) Then
            Set CreateVariableEntry = NewVariableEntry(18, (Addr - 1) * 16, VAR_SIZE_WORD)
        End If
    Case "NAI"
        If ValidateNumber(Addr, 1, 32) Then
            Set CreateVariableEntry = NewVariableEntry(21, (Addr - 1) * 16, VAR_SIZE_WORD)
        End If
    Case "AQ"
        If ValidateNumber(Addr, 1, 8) Then
            Set CreateVariableEntry = NewVariableEntry(19, (Addr - 1) * 16, VAR_SIZE_WORD)
        End If
    Case "NAQ"
        If ValidateNumber(Addr, 1, 16) Then
            Set CreateVariableEntry = NewVariableEntry(22, (Addr - 1) * 16, VAR_SIZE_WORD)
        End If
    Case "AM"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(20, (Addr - 1) * 16, VAR_SIZE_WORD)
        End If
    Case "CUR"
        If ValidateNumber(Addr, 1, 4) Then
            Set CreateVariableEntry = NewVariableEntry(12, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case "F"
        If ValidateNumber(Addr, 1, 4) Then
            Set CreateVariableEntry = NewVariableEntry(13, (Addr - 1), VAR_SIZE_BIT)
        End If
    Case Else
    End Select
End Function
