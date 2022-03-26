VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrendChartIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function GetName(shapeObj)
    GetName = "" ' By default, it shall return empty string to indicate error
    
    If shapeObj.GroupItems.Count <> 7 Then
        Exit Function
    End If
    
    If Left(shapeObj.name, 10) <> "TrendChart" Then
        Exit Function
    End If

    GetName = shapeObj.name
End Function

Public Function GetGadget(shapeObj, sheet As Worksheet, Workbook As Workbook)
    Dim ch As TrendChart
    Set ch = New TrendChart
        
    ch.LoadFromShape shapeObj, sheet, Workbook
    
    Set GetGadget = ch
End Function

Public Function CreateGadget()
    Dim chart As TrendChart
    
    Set chart = New TrendChart
    chart.createNew

    Set CreateGadget = chart
End Function
