VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VariableDataV8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_ranges

Private Sub Class_Initialize()
    Set m_ranges = CreateObject("Scripting.Dictionary")
End Sub

Public Sub FillData(id, ascData, isAscData)
    If ascData <> "" Then
        Dim val
        
        val = CDbl(ascData)
    
        If Not m_ranges.exists(id) Then
            m_ranges.Add id, val
        End If
    End If
End Sub

Public Function GetValue(entry)
    Dim rangeObj
    
    If Not m_ranges.exists(entry.GetReqId()) Then
        GetValue = Empty
    Else
        GetValue = m_ranges.item(entry.GetReqId())
    End If
End Function
