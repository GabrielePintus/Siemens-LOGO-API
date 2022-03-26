VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GadgetManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Identifier is required to provide the following interfaces:
'     . GetName(shapeObj)
'     . GetGadget(shapeObj)
'     . CreateGadget()
' Gadget is required to provide the following interfaces:
'     . GetName()
'     . ChangeName
'     . IsShape(shapeObj)
'     . GetCheckedFlag
'     . SetCheckedFlag

Dim m_identifier ' to get name of a groupshape
Dim m_gadgets
Dim m_WorkBook As Workbook

Public Sub Initialize(identifier, Workbook As Workbook)
    Set m_identifier = identifier

    Set m_gadgets = CreateObject("Scripting.Dictionary")
    
    Set m_WorkBook = Workbook
    
    Update
End Sub

Public Function GetGadgets()
    Set GetGadgets = m_gadgets
End Function

Private Sub HandleGroupShape(shapeObj, sheet As Worksheet)
    Dim name
    name = m_identifier.GetName(shapeObj)
    If name <> "" Then
        ' Only try when it is identified as a valid gadget
        Dim gadget
        
        If m_gadgets.exists(name) Then
            ' It may exists
            If m_gadgets.item(name).IsShape(shapeObj) Then
                Set gadget = m_gadgets.item(name)
            Else
                ' It is another gadget with the same name. The original one may still exists
                ' grand the gadget a new name
                Set gadget = m_identifier.GetGadget(shapeObj, sheet, m_WorkBook)
                
                gadget.ChangeName ' change its name
                
                m_gadgets.Add gadget.GetName, gadget
            End If
        Else
            ' It is a new Shape.
            Set gadget = m_identifier.GetGadget(shapeObj, sheet, m_WorkBook)
        
            m_gadgets.Add gadget.GetName, gadget
        End If
        
        gadget.SetCheckedFlag (True) ' Mark the gadget as checked
    End If
End Sub

' It is possible that Update will fail when user close the workbook while it is updating
' Catch the error and let caller handle the case
Public Function Update()
    On Error GoTo Err ' Probably Close Event Handling

    Dim entries
    entries = m_gadgets.items

    Dim i
    
    ' Mark everyone as Not Checked
    For i = 0 To m_gadgets.Count - 1
        entries(i).SetCheckedFlag (False)
    Next
    
    ' Traverse to check each gadget
    Dim shs As Sheets
    Set shs = m_WorkBook.Sheets ' Application.Sheets may access xlam's sheets
    
    Dim shapeObj
    For i = 1 To shs.Count
        Dim j
        For j = 1 To shs(i).Shapes.Count
            Set shapeObj = shs(i).Shapes.item(j)
            
            If TypeName(shapeObj.OLEFormat.Object) = "GroupObject" Then
                HandleGroupShape shapeObj, shs(i)
            End If
        Next
    Next
    
    ' Remove Those Obsolete ones
    entries = m_gadgets.items
    
    Dim keys
    keys = m_gadgets.keys
    
    Dim totalCount
    totalCount = m_gadgets.Count
    For i = 0 To totalCount - 1
        If Not entries(i).GetCheckedFlag() Then
            m_gadgets.Remove keys(i)
        End If
    Next
    
    Update = True
    Exit Function
    
Err:
    Update = False
End Function

Public Sub HandleEvent(gadgetId, eventId)
    If gadgetId <> "" Then
        If m_gadgets.exists(gadgetId) Then
            m_gadgets.item(gadgetId).HandleEvent eventId
        End If
    Else
        Select Case eventId
        Case "new"
            Create
        End Select
    End If
End Sub

Public Function Create()
    Set gadget = m_identifier.CreateGadget()

    m_gadgets.Add gadget.GetName, gadget
End Function
