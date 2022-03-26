VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogInGadget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim newGroupBox
Dim newTextIp
Dim newTextPwd
Dim newInputIp
Dim newInputPwd
Dim newButton
Dim newBtnBack
Dim theGroup
Dim theGroupName
Dim checkedFlag As Boolean
Dim ownerSheet As Worksheet
Dim ownerWorkbook As Workbook
Dim m_style As Integer

Private Sub Class_Terminate()
    ' When delete a login-owner, Log Out!
    If m_style = LOGIN_GADGET_STATE_CONNECTING Or m_style = LOGIN_GADGET_STATE_CONNECTED Then
        LogOut
    End If
End Sub

' The interface to create new component
Public Function createNew()
    Dim cursorX
    Dim cursorY
    Dim groupId
    
    cursorX = Selection.Left
    cursorY = Selection.Top
    
    ' The ownerSheet
    Set ownerSheet = ActiveSheet
    Set ownerWorkbook = ActiveWorkbook
    
    theGroupName = GetUnusedName(GADGET_TYPE_LOGIN, ownerWorkbook)
    
    Set newBtnBack = ActiveSheet.Buttons.Add(cursorX, cursorY, 220, 128)
    newBtnBack.Characters.Text = ""
    
    Set newGroupBox = ActiveSheet.GroupBoxes.Add(cursorX + 10, cursorY + 10, 200, 108)
    newGroupBox.Characters.Text = "Login Panel (Offline)"
    
    Set newTextIp = ActiveSheet.Labels.Add(cursorX + 25, cursorY + 25, 60, 18)
    newTextIp.Characters.Text = "IP Address"
    
    Set newInputIp = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
        Left:=cursorX + 85, Top:=cursorY + 25, Width:=110, Height:=18)
    newInputIp.Object.MaxLength = 15
        
    Set newTextPwd = ActiveSheet.Labels.Add(cursorX + 25, cursorY + 55, 60, 18)
    newTextPwd.Characters.Text = "Password"
    
    Set newInputPwd = ActiveSheet.OLEObjects.Add(ClassType:="Forms.TextBox.1", Link:=False, DisplayAsIcon:=False, _
        Left:=cursorX + 85, Top:=cursorY + 55, Width:=110, Height:=18)
    newInputPwd.Object.PasswordChar = "*"
    newInputPwd.Object.MaxLength = 10
    
    Set newButton = ActiveSheet.Buttons.Add(cursorX + 145, cursorY + 85, 50, 18)
    newButton.Characters.Text = "Log In"
    
    m_style = LOGIN_GADGET_STATE_NORMAL

    UpdateOnActionHandler
    
    GroupComponent
End Function

Private Sub UpdateOnActionHandler()
    newButton.OnAction = "'ActiveWorkBookEvent """ + GADGET_TYPE_LOGIN + """,""" + theGroupName + """,""login""'"
End Sub

' The interface to load component from shape object
Public Function LoadFromShape(groupShape, sheet As Worksheet, Workbook As Workbook)
    theGroupName = groupShape.name

    Set newBtnBack = groupShape.GroupItems.item(1).OLEFormat.Object
    Set newGroupBox = groupShape.GroupItems.item(2).OLEFormat.Object
    Set newTextIp = groupShape.GroupItems.item(3).OLEFormat.Object
    Set newInputIp = groupShape.GroupItems.item(4).OLEFormat.Object
    Set newTextPwd = groupShape.GroupItems.item(5).OLEFormat.Object
    Set newInputPwd = groupShape.GroupItems.item(6).OLEFormat.Object
    Set newButton = groupShape.GroupItems.item(7).OLEFormat.Object
    
    Set theGroup = groupShape
    
    Set ownerSheet = sheet
    Set ownerWorkbook = Workbook
    
    ' Initialize the gadget into normal state after syncing
    SetStyle LOGIN_GADGET_STATE_NORMAL
End Function

Public Sub SetStyle(style)
    m_style = style

    ' Simple Optimization, Don't bother if it is definitely OK
    If newInputIp.Object.Enabled And style = LOGIN_GADGET_STATE_NORMAL Then
        Exit Sub
    End If

    UnGroupComponent
    
    Select Case style
    Case LOGIN_GADGET_STATE_NORMAL
        newGroupBox.Characters.Text = "Login Panel (Offline)"
        newInputIp.Object.Enabled = True
        newInputPwd.Object.Enabled = True
        newButton.Characters.Text = "Log In"
        newButton.Enabled = True
        newButton.Characters.Font.ColorIndex = xlColorIndexAutomatic
    Case LOGIN_GADGET_STATE_CONNECTING
        newGroupBox.Characters.Text = "Login Panel (Connecting)"
        newInputIp.Object.Enabled = False
        newInputPwd.Object.Enabled = False
        newButton.Characters.Text = "Log In ..."
        newButton.Enabled = False
        newButton.Characters.Font.ColorIndex = 15 ' Gray
    Case LOGIN_GADGET_STATE_CONNECTED
        newGroupBox.Characters.Text = "Login Panel (Online)"
        newInputIp.Object.Enabled = False
        newInputPwd.Object.Enabled = False
        newButton.Characters.Text = "Log Out"
        newButton.Enabled = True
        newButton.Characters.Font.ColorIndex = xlColorIndexAutomatic
    Case LOGIN_GADGET_STATE_DISABLED
        newGroupBox.Characters.Text = "Login Panel (Offline)"
        newInputIp.Object.Enabled = False
        newInputPwd.Object.Enabled = False
        newButton.Characters.Text = "Log In"
        newButton.Enabled = False
        newButton.Characters.Font.ColorIndex = 15 ' Gray
    Case Else
        ' Do nothing if it is an invalid style
    End Select
  
    GroupComponent
End Sub

Public Sub SetStatus(status)
    If m_style = LOGIN_GADGET_STATE_CONNECTED Then
        UnGroupComponent
            Select Case status
            Case 0:
                newGroupBox.Characters.Text = "Login Panel (Online)"
            Case 1:
                newGroupBox.Characters.Text = "Login Panel (Online - Stop)"
            Case 2:
                newGroupBox.Characters.Text = "Login Panel (Online - Running)"
            End Select
        GroupComponent
    End If
End Sub

Public Function GetName()
    GetName = theGroupName
End Function

Public Sub ChangeName()
    theGroupName = GetUnusedName(GADGET_TYPE_LOGIN, ownerWorkbook)
    
    UnGroupComponent
    
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
    Set theGroup = ownerSheet.Shapes.Range(Array(newButton.name, newGroupBox.name, newTextIp.name, newInputIp.name, _
        newTextPwd.name, newInputPwd.name, newBtnBack.name)).group
        
    theGroup.name = theGroupName
End Sub

Public Sub HandleEvent(eventId)
    ' GetConnectionState() ' another solution is determined by Connection State
    If newInputIp.Object.Enabled Then
        ' It is definitely in Not_Connected state
        
        Dim ipValue
        Dim pwdValue
    
        ipValue = newInputIp.Object.value
        pwdValue = newInputPwd.Object.value
        
        ' Validate the IP address
        If Not GetIPValidateReg(ipValue) Then
            MsgBox "Invalid Ip Address"
            Exit Sub
        End If
        
        ' Perform Gadget LogIn Then
        GadgetLogIn ipValue, pwdValue, Me.GetName
    Else
        ' It is in state other than Not_Connected
        GadgetLogOut Me.GetName
    End If
End Sub
