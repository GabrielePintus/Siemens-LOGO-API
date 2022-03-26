VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogInGadgetIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetName(shapeObj)
    GetName = "" ' By default, it shall return empty string to indicate error
    
    If shapeObj.GroupItems.Count <> 7 Then
        Exit Function
    End If
    
    Dim btn
    Set btn = shapeObj.GroupItems.item(7).OLEFormat.Object
    
    If TypeName(btn) <> "Button" Then
        Exit Function
    End If
    
    'Dim btnAction
    'btnAction = btn.OnAction ' onAction failed to be get in grouped mode

    If Left(shapeObj.name, 5) <> "LogIn" Then
        Exit Function
    End If

    GetName = shapeObj.name
End Function

Public Function GetGadget(shapeObj, sheet As Worksheet, Workbook As Workbook)
    Dim obj As LogInGadget
    Set obj = New LogInGadget
        
    obj.LoadFromShape shapeObj, sheet, Workbook
    Set GetGadget = obj
End Function

Public Function CreateGadget()
    Dim obj As LogInGadget
    
    Set obj = New LogInGadget
    obj.createNew
    
    Set CreateGadget = obj
End Function
