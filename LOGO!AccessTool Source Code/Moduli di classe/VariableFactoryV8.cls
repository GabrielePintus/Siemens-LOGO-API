VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariableFactoryV8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Create2DVariableEntry(Range As String, Addr As Integer, subAddr As Integer)
    Set Create2DVariableEntry = Nothing ' Return Nothing by default to indicate error
    
    Select Case Range
    Case "SR"
        If ValidateNumber(Addr, 1, 4) And ValidateNumber(subAddr, 1, 8) Then
            Set Create2DVariableEntry = NewVariableEntry(10, (Addr - 1) * 8 + subAddr, VAR_SIZE_BIT)
        End If
    Case Else
    End Select
End Function

Public Function CreateVariableEntry(Range As String, Addr As Integer)
    Set CreateVariableEntry = Nothing ' Return Nothing by default to indicate error
    
    Select Case Range
    Case "VB"
        If ValidateNumber(Addr, 0, 850) Then
            Set CreateVariableEntry = NewVariableEntry(21, Addr, VAR_SIZE_BYTE)
        End If
    Case "VW"
        If ValidateNumber(Addr, 0, 849) Then
            Set CreateVariableEntry = NewVariableEntry(21, Addr, VAR_SIZE_WORD)
        End If
    Case "VD"
        If ValidateNumber(Addr, 0, 847) Then
            Set CreateVariableEntry = NewVariableEntry(21, Addr, VAR_SIZE_DWORD)
        End If
    Case "I"
        If ValidateNumber(Addr, 1, 24) Then
            Set CreateVariableEntry = NewVariableEntry(7, Addr, VAR_SIZE_BIT)
        End If
    Case "NI"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(11, Addr, VAR_SIZE_BIT)
        End If
    Case "Q"
        If ValidateNumber(Addr, 1, 20) Then
            Set CreateVariableEntry = NewVariableEntry(5, Addr, VAR_SIZE_BIT)
        End If
    Case "NQ"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(17, Addr, VAR_SIZE_BIT)
        End If
    Case "M"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(3, Addr, VAR_SIZE_BIT)
        End If
    Case "AI"
        If ValidateNumber(Addr, 1, 8) Then
            Set CreateVariableEntry = NewVariableEntry(14, Addr, VAR_SIZE_WORD)
        End If
    Case "NAI"
        If ValidateNumber(Addr, 1, 32) Then
            Set CreateVariableEntry = NewVariableEntry(15, Addr, VAR_SIZE_WORD)
        End If
    Case "AQ"
        If ValidateNumber(Addr, 1, 8) Then
            Set CreateVariableEntry = NewVariableEntry(6, Addr, VAR_SIZE_WORD)
        End If
    Case "NAQ"
        If ValidateNumber(Addr, 1, 16) Then
            Set CreateVariableEntry = NewVariableEntry(18, Addr, VAR_SIZE_WORD)
        End If
    Case "AM"
        If ValidateNumber(Addr, 1, 64) Then
            Set CreateVariableEntry = NewVariableEntry(4, Addr, VAR_SIZE_WORD)
        End If
    Case "CUR"
        If ValidateNumber(Addr, 1, 4) Then
            Set CreateVariableEntry = NewVariableEntry(8, Addr, VAR_SIZE_BIT)
        End If
    Case "F"
        If ValidateNumber(Addr, 1, 4) Then
            Set CreateVariableEntry = NewVariableEntry(9, Addr, VAR_SIZE_BIT)
        End If
    Case Else
    End Select
End Function
