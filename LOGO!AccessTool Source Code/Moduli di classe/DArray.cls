VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_data()
Dim m_count As Integer

Private Sub Class_Initialize()
    ReDim m_data(256)
End Sub

Public Sub Add(item As Object)
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
