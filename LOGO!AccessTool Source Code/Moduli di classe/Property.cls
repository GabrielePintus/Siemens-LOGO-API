VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Property"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_id
Dim m_value
Dim m_propertyListeners

Private Sub Class_Initialize()
    Set m_propertyListeners = CreateObject("Scripting.Dictionary")
End Sub

Public Sub SetId(id)
    m_id = id
End Sub

Public Sub SetValue(newValue, arg0)
    If newValue <> m_value Then
        DebugLog "[Property] " + m_id + ": " + CStr(m_value) + " > " + CStr(newValue)
    
        m_value = newValue
        
        NotifyAllListeners newValue, arg0
        
        RefreshVolatile
    End If
End Sub

Public Function GetValue()
    GetValue = m_value
End Function

Public Sub AddListener(action)
    If Not m_propertyListeners.exists(action) Then
        m_propertyListeners.Add action, False ' mark its default
        
        ' trigger connected notification if it is already in such a state
    End If
End Sub

Public Sub RemoveListener(action)
    If m_propertyListeners.exists(action) Then
        m_propertyListeners.Remove action
        
        ' trigger connected notification if it is already in such a state
    End If
End Sub

Private Sub NotifyAllListeners(newValue, arg0)
    Dim actions
    Dim i

    actions = m_propertyListeners.keys

    For i = 0 To m_propertyListeners.Count - 1
        NotifyListener actions(i), newValue, arg0
    Next
End Sub

Private Sub NotifyListener(action, newValue, arg0)
    Application.Run action, newValue, arg0
End Sub


