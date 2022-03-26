VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrendChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' group name
Dim theGroupName As String

' components
Dim chartObj
Dim newBtnBack
Dim variableGroupBox
Dim variableInput
Dim addButton
Dim removeButton
Dim variableList

' the owner which contains the Trend Chart
Dim ownerSheet As Worksheet
Dim ownerWorkbook As Workbook

' the grouped shape
Dim theGroup

' For convenient opeartion of data sheet of chartobject
Dim dataSheet

' Checked Flag
Dim checkedFlag As Boolean

Const TREND_CHART_DATA_SIZE As Integer = 50

Public Sub createNew()
    Dim cursorX
    Dim cursorY

    cursorX = Selection.Left
    cursorY = Selection.Top
    
    ' The ownerSheet
    Set ownerSheet = ActiveSheet
    Set ownerWorkbook = ActiveWorkbook
    
    ' The group name
    theGroupName = GetUnusedName(GADGET_TYPE_TRENDCHART, ownerWorkbook)
    
    ' The back for control panel
    Set newBtnBack = ActiveSheet.Buttons.Add(cursorX + 0, cursorY + 0, 520, 235)
    newBtnBack.Characters.Text = ""
    
    ' The chart object
    Set chartObj = ActiveSheet.OLEObjects.Add(ClassType:="MSGraph.Chart.8", Link:=True, DisplayAsIcon:=False, _
        Left:=cursorX + 6, Top:=cursorY + 6, Width:=354, Height:=224)
    Set dataSheet = chartObj.Object.Application.dataSheet
    
    InitializeChart ' Call function to initialize the chart object (and dataSheet)
    
    ' Other components of control panel
    Set variableGroupBox = ActiveSheet.GroupBoxes.Add(cursorX + 365, cursorY + 10, 150, 220)
    variableGroupBox.Characters.Text = "Variable Settings"
    
    Set variableInput = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
        Left:=cursorX + 371, Top:=cursorY + 21, Width:=54, Height:=18)
    
    Set addButton = ActiveSheet.Buttons.Add(cursorX + 429, cursorY + 21, 28, 18)
    addButton.Characters.Text = "Add"
    
    Set removeButton = ActiveSheet.Buttons.Add(cursorX + 461, cursorY + 21, 48, 18)
    removeButton.Characters.Text = "Remove"
    
    Set variableList = ActiveSheet.ListBoxes.Add(cursorX + 371, cursorY + 45, 138, 179)
    variableList.Display3DShading = True
    
    'Set variableList = ActiveSheet.OLEObjects.Add(ClassType:="Forms.ListBox.1", Link:=False, _
    '    DisplayAsIcon:=False, Left:=cursorX + 371, Top:=cursorY + 45, Width:=138, Height:=134)
    'variableList.AutoLoad = True
    
    ' bind operations
    UpdateOnActionHandler
    
    ' group components
    GroupComponent
End Sub

Private Sub UpdateOnActionHandler()
    addButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_TRENDCHART + """,""" + theGroupName + """,""add""'"
    removeButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_TRENDCHART + """,""" + theGroupName + """,""remove""'"
End Sub

Private Sub InitializeChart()
    chartObj.AutoLoad = True
    chartObj.Object.ChartType = 4 ' make it pure line
    
    ' Change Style Before Clearing Data (Otherwise, it will fail when it is empty)
    
    ' Change Style of Legend
    With chartObj.Object.Legend.Font
        .Bold = True
        .size = 8
        .ColorIndex = xlColorIndexAutomatic
        .name = "Arial"
    End With

    ' Change Style of Y Axis
    With chartObj.Object.Axes(xlValue).TickLabels.Font
        .Bold = False
        .size = 8
        .ColorIndex = xlColorIndexAutomatic
        .name = "Arial"
    End With
    
    ' Change Style of X Axis
    With chartObj.Object.Axes(xlCategory).TickLabels.Font
        .Bold = False
        .size = 8
        .ColorIndex = xlColorIndexAutomatic
        .name = "Arial"
    End With
    
    'Remove all 4 lines of the datasheet. they are automatically filled when created
    dataSheet.Rows(1).Delete
    dataSheet.Rows(1).Delete
    dataSheet.Rows(1).Delete
    dataSheet.Rows(1).Delete
End Sub

' The interface to load component from shape object
Public Function LoadFromShape(groupShape, sheet As Worksheet, Workbook As Workbook)
    theGroupName = groupShape.name

    Set newBtnBack = groupShape.GroupItems.item(1).OLEFormat.Object
    Set chartObj = groupShape.GroupItems.item(2).OLEFormat.Object
    Set variableGroupBox = groupShape.GroupItems.item(3).OLEFormat.Object
    Set variableInput = groupShape.GroupItems.item(4).OLEFormat.Object
    Set addButton = groupShape.GroupItems.item(5).OLEFormat.Object
    Set removeButton = groupShape.GroupItems.item(6).OLEFormat.Object
    Set variableList = groupShape.GroupItems.item(7).OLEFormat.Object
    
    ' Remarks: This property of ChartObject will not be copied. Forcely override it!
    chartObj.AutoLoad = True
    
    Set theGroup = groupShape
    
    ' The ownerSheet
    Set ownerSheet = sheet
    Set ownerWorkbook = Workbook
    
    Set dataSheet = chartObj.Object.Application.dataSheet
End Function

Public Function GetName()
    GetName = theGroupName
End Function

Public Sub ChangeName()
    theGroupName = GetUnusedName(GADGET_TYPE_TRENDCHART, ownerWorkbook)
    
    UnGroupComponent
    
    ' bind operations
    UpdateOnActionHandler
    
    GroupComponent
End Sub

Public Function GetCheckedFlag()
    GetCheckedFlag = checkedFlag
End Function

Public Sub SetCheckedFlag(flag As Boolean)
    checkedFlag = flag
End Sub

Public Function IsShape(groupShape)
    If groupShape.id = theGroup.id Then
        IsShape = True
    Else
        IsShape = False
    End If
End Function

Private Sub UnGroupComponent()
    theGroup.Ungroup
End Sub

Private Sub GroupComponent()
    Set theGroup = ownerSheet.Shapes.Range(Array(chartObj.name, newBtnBack.name, variableGroupBox.name, variableInput.name, addButton.name, _
        removeButton.name, variableList.name)).group
       
    theGroup.name = theGroupName
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Button Handler
Public Sub HandleEvent(eventId)
    Select Case eventId
    Case "add"
        AddVariable
    Case "remove"
        RemoveVariable
    End Select
End Sub


Public Sub AddVariable()
    Dim inputtedId
    inputtedId = UCase(variableInput.Object.value)
    
    If inputtedId = "" Then
        Exit Sub
    End If
    
    ' Clear inputted value no matter what it was
    variableInput.Object.value = ""

    Dim entry
    Set entry = GetVariableEntry(inputtedId, False) ' Don't add invalid one because it is just a test
    
    ' Do nothing if it is invalid
    If entry Is Nothing Then
        MsgBox "Invalid Variable ID"
        Exit Sub
    End If
    
    UnGroupComponent
    
    variableList.AddItem entry.GetReqId()
    
    Dim variableIndex
    variableIndex = variableList.ListCount
    
    GroupComponent
    
    ' Add this variable in Chart
    dataSheet.Cells(variableIndex + 1, 1) = entry.GetReqId()
End Sub

Public Sub RemoveVariable()
    UnGroupComponent
    
    Dim selectedIndex
    selectedIndex = variableList.ListIndex
    
    If selectedIndex > 0 Then
        variableList.RemoveItem selectedIndex
        
        Dim Count
        Count = variableList.ListCount
        
        If selectedIndex <= Count Then
            variableList.ListIndex = selectedIndex
        Else
            If Count > 0 Then
                variableList.ListIndex = Count
            End If
        End If
    End If
    
    GroupComponent
    
    If selectedIndex > 0 Then
        ' Remove this variable in Chart
        dataSheet.Rows(selectedIndex + 1).Delete
    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interface to Refresh Data

Public Sub RefreshData(timeStr, dataLogTimeStr)
    On Error GoTo Err ' Probably Close Event Handling

    ' Refresh Data Line (if there is any)
    If Not IsNull(variableList.List) Then
        dataSheet.Columns(2).Delete

        ' Add time label
        dataSheet.Cells(1, TREND_CHART_DATA_SIZE) = timeStr
    
        Dim i
        For i = LBound(variableList.List) To UBound(variableList.List)
            Dim var
            var = GetVariableValue(variableList.List(i))
            
            ' fill it in datasheet
            dataSheet.Cells(i + 1, TREND_CHART_DATA_SIZE) = var
        Next
    End If
    
    Exit Sub
    
Err:
    ' Simply do nothing: If it is caused by close, it will be recycled by manager.
    ' If it is caused by something else, its misoperation will also be detected as
    ' a bug.
End Sub

Public Sub StopRefreshing()
End Sub

