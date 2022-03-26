Attribute VB_Name = "PropertyManager"
Option Private Module
Dim m_properties As Object

Public Const PROPERTY_ID_STATUS As String = "status"


Public Const PROPERTY_ID_CONNECTION As String = "connection"
Public Const STATE_NOT_CONNECTED = 0
Public Const STATE_CONNECTING = 1
Public Const STATE_CONNECTED = 2
Public Const STATE_RECOVERING = 3

Private Function GetPropertyObject(propertyId)
    If m_properties Is Nothing Then
        Set m_properties = CreateObject("Scripting.Dictionary")
    End If
    
    If Not m_properties.exists(propertyId) Then
        Dim newPropertyObj
        Set newPropertyObj = New Property
        
        newPropertyObj.SetId propertyId
    
        m_properties.Add propertyId, newPropertyObj
    End If
    
    Set GetPropertyObject = m_properties.item(propertyId)
End Function

Public Sub SetProperty(propertyId, newValue, arg0)
    GetPropertyObject(propertyId).SetValue newValue, arg0
End Sub

Public Function GetProperty(propertyId)
    GetProperty = GetPropertyObject(propertyId).GetValue
End Function

Public Sub AddPropertyListener(propertyId, action)
    GetPropertyObject(propertyId).AddListener action
End Sub

Public Sub RemovePropertyListener(propertyId, action)
    GetPropertyObject(propertyId).RemoveListener action
End Sub
