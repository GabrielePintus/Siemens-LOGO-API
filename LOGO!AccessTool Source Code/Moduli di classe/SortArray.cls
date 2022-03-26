VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_data()
Dim m_count As Integer

Private Sub Class_Initialize()
    ReDim m_data(256)
End Sub

Public Sub Add(id, Column, entry As Object)
    ' create a new item
    Dim item As SortItem
    Set item = New SortItem
    
    item.m_id = id
    item.m_Column = Column
    Set item.m_variableEntry = entry

    If m_count > UBound(m_data) Then
        ReDim Preserve m_data(UBound(m_data) * 2)
    End If
    
    Set m_data(m_count) = item
    m_count = m_count + 1
End Sub

Public Function Count()
    Count = m_count
End Function

Public Function GetItem(index)
    Set GetItem = m_data(index)
End Function

Public Sub Sort()
    SortPart 0, m_count
End Sub

Public Sub Clear()
    Dim i
    For i = 0 To m_count - 1
        Set m_data(i) = Nothing
    Next
    
    m_count = 0
End Sub

Public Sub Test()
    Dim STR As String
    
    Dim i
    For i = 0 To m_count - 1
        STR = STR + "," + CStr(m_data(i))
    Next

    MsgBox STR
End Sub

Public Function GetHead()
    Dim STR As String
    
    Dim i
    For i = 0 To m_count - 1
        STR = STR + "," + CStr(m_data(i).m_id)
    Next
    
    GetHead = STR
End Function

Public Function GetData()
    Dim STR As String
    
    Dim i
    For i = 0 To m_count - 1
        Dim entry As Object
        Set entry = m_data(i).m_variableEntry
        
        If entry Is Nothing Then
            STR = STR + "," + GetStatusString()
        Else
            STR = STR + "," + CStr(entry.GetValue())
        End If
    Next

    GetData = STR
End Function

Private Function CompareItem(p0, p1)
    Dim item0
    Dim item1
    
    Set item0 = m_data(p0)
    Set item1 = m_data(p1)

    Dim columnDiff
    columnDiff = item0.m_Column - item1.m_Column
    
    ' Check if column is different
    If columnDiff > 0 Then
        CompareItem = 1
        Exit Function
    Else
        If columnDiff < 0 Then
            CompareItem = -1
            Exit Function
        End If
    End If
    
    ' Comes here implies they have the same column
    Dim entry0 As VariableEntry
    Dim entry1 As VariableEntry
    
    Set entry0 = item0.m_variableEntry
    Set entry1 = item1.m_variableEntry
    
    ' Check if they are status which has no entry
    If entry0 Is Nothing Then
        If entry1 Is Nothing Then
            CompareItem = 0
            Exit Function
        Else
            CompareItem = -1
            Exit Function
        End If
    Else
        If entry1 Is Nothing Then
            CompareItem = 1
            Exit Function
        End If
    End If
    
    ' Comes here implies they both have entry
    Dim diff
    diff = entry0.m_Range - entry1.m_Range
    
    If diff <> 0 Then
        CompareItem = diff
        Exit Function
    End If
    
    diff = entry0.m_addr - entry1.m_addr
    
    If diff <> 0 Then
        CompareItem = diff
        Exit Function
    End If
    
    diff = entry0.m_size - entry1.m_size
    
    If diff <> 0 Then
        CompareItem = diff
        Exit Function
    End If
    
    CompareItem = 0
End Function

Private Sub Switch(p0, p1)
    Dim item
    
    Set item = m_data(p0)
    Set m_data(p0) = m_data(p1)
    Set m_data(p1) = item
End Sub

Private Sub SortPart(a0, a1)
    Dim i
    Dim j
    i = a0
    j = a1 - 1
    
    Do While j > i
        ' keep searching a smaller one from end until it meets i
        While CompareItem(j, a0) >= 0 And j > i
            j = j - 1
        Wend
    
        ' keep searching a bigger one from start until it meets j
        While CompareItem(i, a0) <= 0 And i < j
            i = i + 1
        Wend
    
        If i <> j Then
            ' switch those two items if they are not the same one
            Switch i, j
        Else
            ' the meet point must be a equal(lesser) of base
            If i <> a0 Then
               Switch i, a0
               
               SortPart a0, i ' recursively sort its left wing
            End If
            
            SortPart i + 1, a1 ' recursively sort its right wing
        End If
    Loop
End Sub

