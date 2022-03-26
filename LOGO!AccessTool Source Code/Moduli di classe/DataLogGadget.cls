VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLogGadget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' group name
Dim theGroupName As String

' components
Dim newBtnBack
Dim variableGroupBox
Dim variableInput
Dim addButton
Dim removeButton
Dim variableList
Dim dataLogGroupBox
Dim dataLogDirInput
Dim browseButton

' the owner which contains the Trend Chart
Dim ownerSheet As Worksheet
Dim ownerWorkbook As Workbook

' the grouped shape
Dim theGroup

' Data Log Handle
Dim LOG As DataLog

' Checked Flag
Dim checkedFlag As Boolean

Const TREND_CHART_DATA_SIZE As Integer = 50

Private Sub Class_Initialize()
    Set LOG = New DataLog
End Sub

Public Sub createNew()
    Dim cursorX
    Dim cursorY

    cursorX = Selection.Left
    cursorY = Selection.Top
    
    ' The ownerSheet
    Set ownerSheet = ActiveSheet
    Set ownerWorkbook = ActiveWorkbook
    
    ' The group name
    theGroupName = GetUnusedName(GADGET_TYPE_DATALOG, ownerWorkbook)
    
    ' The back for control panel
    Set newBtnBack = ActiveSheet.Buttons.Add(cursorX + 0, cursorY + 0, 161, 235)
    newBtnBack.Characters.Text = ""
    
    ' Other components of control panel
    Set variableGroupBox = ActiveSheet.GroupBoxes.Add(cursorX + 6, cursorY + 10, 150, 220)
    variableGroupBox.Characters.Text = "Data Log"
    
    Set variableInput = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
        Left:=cursorX + 12, Top:=cursorY + 21, Width:=54, Height:=18)
    
    Set addButton = ActiveSheet.Buttons.Add(cursorX + 70, cursorY + 21, 28, 18)
    addButton.Characters.Text = "Add"
    
    Set removeButton = ActiveSheet.Buttons.Add(cursorX + 102, cursorY + 21, 48, 18)
    removeButton.Characters.Text = "Remove"
    
    Set variableList = ActiveSheet.ListBoxes.Add(cursorX + 12, cursorY + 45, 138, 140)
    variableList.Display3DShading = True
    
    'Set variableList = ActiveSheet.OLEObjects.Add(ClassType:="Forms.ListBox.1", Link:=False, _
    '    DisplayAsIcon:=False, Left:=cursorX + 371, Top:=cursorY + 45, Width:=138, Height:=134)
    'variableList.AutoLoad = True
    
    Set dataLogGroupBox = ActiveSheet.Labels.Add(cursorX + 12, cursorY + 190, 100, 18)
    dataLogGroupBox.Characters.Text = "Data Log Directory"
    
    Set dataLogDirInput = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
        Left:=cursorX + 12, Top:=cursorY + 206, Width:=89, Height:=18)
    
    Set browseButton = ActiveSheet.Buttons.Add(cursorX + 105, cursorY + 206, 45, 18)
    browseButton.Characters.Text = "Browse"
    
    ' bind operations
    UpdateOnActionHandler
    
    ' group components
    GroupComponent
End Sub

Private Sub UpdateOnActionHandler()
    addButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_DATALOG + """,""" + theGroupName + """,""add""'"
    removeButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_DATALOG + """,""" + theGroupName + """,""remove""'"
    browseButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_DATALOG + """,""" + theGroupName + """,""browse""'"
End Sub

' The interface to load component from shape object
Public Function LoadFromShape(groupShape, sheet As Worksheet, Workbook As Workbook)
    theGroupName = groupShape.name

    Set newBtnBack = groupShape.GroupItems.item(1).OLEFormat.Object
    Set variableGroupBox = groupShape.GroupItems.item(2).OLEFormat.Object
    Set variableInput = groupShape.GroupItems.item(3).OLEFormat.Object
    Set addButton = groupShape.GroupItems.item(4).OLEFormat.Object
    Set removeButton = groupShape.GroupItems.item(5).OLEFormat.Object
    Set variableList = groupShape.GroupItems.item(6).OLEFormat.Object
    Set dataLogGroupBox = groupShape.GroupItems.item(7).OLEFormat.Object
    Set dataLogDirInput = groupShape.GroupItems.item(8).OLEFormat.Object
    Set browseButton = groupShape.GroupItems.item(9).OLEFormat.Object

    Set theGroup = groupShape
    
    ' The ownerSheet
    Set ownerSheet = sheet
    Set ownerWorkbook = Workbook

    ' load data log dir setting.
    UpdateDataLogDirSetting
    
    ' load variable setting
    UpdateVariableSetting
End Function

Public Function GetName()
    GetName = theGroupName
End Function

Public Sub ChangeName()
    theGroupName = GetUnusedName(GADGET_TYPE_DATALOG, ownerWorkbook)
    
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
    Set theGroup = ownerSheet.Shapes.Range(Array(newBtnBack.name, variableGroupBox.name, variableInput.name, addButton.name, _
        removeButton.name, variableList.name, dataLogGroupBox.name, dataLogDirInput.name, browseButton.name)).group
       
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
    Case "browse"
        BrowseDataLogDir
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
    
    ' Update variable settings
    UpdateVariableSetting
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
        ' Update variable settings
        UpdateVariableSetting
    End If
End Sub

Public Sub BrowseDataLogDir()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' show dialog
    If fd.Show = -1 Then
        ' user did select something
        Dim strPath As String
        
        strPath = fd.SelectedItems(1)
        
        dataLogDirInput.Object.value = strPath
        
        ' Try to load this setting
        UpdateDataLogDirSetting
    End If
    
    Set fd = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interface to Refresh Data

Public Sub RefreshData(timeStr, dataLogTimeStr)
    On Error GoTo Err ' Probably Close Event Handling

    ' Refresh Data Line (if there is any)
    If Not IsNull(variableList.List) Then
        Dim dataStr As String
        
        dataStr = dataLogTimeStr
        
        Dim i
        For i = LBound(variableList.List) To UBound(variableList.List)
            Dim var
            var = GetVariableValue(variableList.List(i))
            
            ' compose dataline
            dataStr = dataStr + "," + CStr(var)
        Next
        
        ' Update Data Log
        LOG.WriteLog dataStr
    End If
    
    Exit Sub
    
Err:
    ' Simply do nothing: If it is caused by close, it will be recycled by manager.
    ' If it is caused by something else, its misoperation will also be detected as
    ' a bug.
End Sub

Public Sub StopRefreshing()
    LOG.CloseDir ' Close Dir when stop refreshing
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Interface to Update DataLog Settings

Private Sub UpdateDataLogDirSetting()
    Dim dirPath As String
    dirPath = dataLogDirInput.Object.value
    
    ' No matter it is empty or not, it shall be set
    
    If Not LOG.OpenDir(dirPath) Then
        MsgBox "Invalid Data Log Directory Path:" + dirPath
    End If
End Sub

Private Sub UpdateVariableSetting()
    Dim headStr As String
    If Not IsNull(variableList.List) Then
        headStr = "Time"
    
        Dim i
        For i = LBound(variableList.List) To UBound(variableList.List)
            headStr = headStr + "," + variableList.List(i)
        Next
    End If

    LOG.SetHead (headStr)
End Sub


