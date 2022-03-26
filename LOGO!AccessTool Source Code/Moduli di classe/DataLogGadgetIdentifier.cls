VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLogGadgetIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetName(shapeObj)
    GetName = "" ' By default, it shall return empty string to indicate error
    
    If shapeObj.GroupItems.Count <> 9 Then
        Exit Function
    End If
    
    If Left(shapeObj.name, 7) <> GADGET_TYPE_DATALOG Then
        Exit Function
    End If

    GetName = shapeObj.name
End Function

Public Function GetGadget(shapeObj, sheet As Worksheet, Workbook As Workbook)
    Dim obj As DataLogGadget
    Set obj = New DataLogGadget
        
    obj.LoadFromShape shapeObj, sheet, Workbook
    
    Set GetGadget = obj
End Function

Public Function CreateGadget()
    Dim obj As DataLogGadget
    
    Set obj = New DataLogGadget
    obj.createNew

    Set CreateGadget = obj
End Function

