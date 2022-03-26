VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkBookContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_WorkBook As Workbook
Dim m_charts As gadgetManager
Dim m_logInGadgets As gadgetManager
Dim m_dataLogGadgets As gadgetManager
Dim m_enableVARL As Boolean

' Data Log Handle
Dim LOG As DataLog
Dim m_array As TypeSortArray

Dim m_gadgetManagers

Private Sub Class_Initialize()
    Set LOG = New DataLog
    Set m_array = New TypeSortArray
    Set m_gadgetManagers = CreateObject("Scripting.Dictionary")
End Sub

Public Sub Load(Workbook As Workbook)
    Set m_WorkBook = Workbook
    
    'AddGadgetManager GADGET_TYPE_DATALOG, New DataLogGadgetIdentifier, Workbook
    'AddGadgetManager GADGET_TYPE_LOGIN, New LogInGadgetIdentifier, Workbook
    'AddGadgetManager GADGET_TYPE_TRENDCHART, New TrendChartIdentifier, Workbook
End Sub

Private Sub AddGadgetManager(id, identifier, Workbook As Workbook)
    If Not m_gadgetManagers.exists(id) Then
        Dim mgr
        Set mgr = New gadgetManager
        
        mgr.Initialize identifier, Workbook
    
        m_gadgetManagers.Add id, mgr
    End If
End Sub

Public Sub StartVARL()
    m_array.Clear
    m_enableVARL = True
End Sub

Public Sub AddVARL(formattedId, Column, entry As Object)
    If m_enableVARL Then
        m_array.Add formattedId, Column, entry ' Record it in DataLog collector
    End If
End Sub

Public Function Update()
    Update = False ' return false by default
    
    Dim entries
    entries = m_gadgetManagers.items

    Dim i
    
    ' Update every gadget managers
    For i = 0 To m_gadgetManagers.Count - 1
        If Not entries(i).Update Then
            Exit Function
        End If
    Next
    
    If Not CheckSheets() Then
        Exit Function
    End If
    
    Update = True
End Function

Public Function CheckSheets()
    CheckSheets = False ' return false by default
    
    On Error GoTo Err ' Probably Close Event Handling
    
    Dim shs As Sheets
    Set shs = m_WorkBook.Sheets ' Application.Sheets may access xlam's sheets
    
    CheckSheets = True
    
    Exit Function
    
Err:
    ' Simply do nothing: If it is caused by close, it will be recycled by manager.
    ' If it is caused by something else, its misoperation will also be detected as
    ' a bug.
End Function

Public Sub CloseVARL(timeStr)
    On Error GoTo Err ' Probably Close Event Handling
   
    ' Disable VARL after saved Log. It will be reopened by start
    m_enableVARL = False
    
    ' Update Data LOG
    Dim varlHead As String
    Dim varlData As String
    
    m_array.Sort
    
    varlHead = m_array.GetHead()
    varlData = m_array.GetData()
    'this is from VariableSyncManager.bas
    If Getdatalogrecord() = 1 Then
        If varlHead <> "" Then
             AppendDataLog "Time" + varlHead, timeStr + varlData
        End If
        Setdatalogrecord (0)
    End If
    
    Exit Sub
    
Err:
    ' Simply do nothing: If it is caused by close, it will be recycled by manager.
    ' If it is caused by something else, its misoperation will also be detected as
    ' a bug.
    MsgBox "error"
End Sub

Public Sub StopRefreshing()
    LOG.CloseDir ' Close Dir when stop refreshing
End Sub

Public Sub AppendDataLog(head As String, data As String)
    Dim fullName
    
    fullName = m_WorkBook.fullName
    
    ' Always try to update head, fileprefix, dir setting because it may changes at any time
    LOG.SetHead head
    LOG.SetFilePrefix GetFileSystemObject().GetBaseName(fullName) + "_" + GetUrl() + "_"
    LOG.OpenDir GetFileSystemObject().GetParentFolderName(fullName)
    
    ' Only try to initialize when it is not successfully opened (file opened)
    If Not LOG.IsSuccess() Then
        LOG.OpenDir Environ("USERPROFILE") + "\Documents"
    End If

    LOG.WriteLog data
End Sub

' Obsolete version of Refresh
'Public Sub Refresh(timeStr)
    'RefreshGadgetManager GADGET_TYPE_DATALOG, timeStr
    'RefreshGadgetManager GADGET_TYPE_TRENDCHART, timeStr
'End Sub

Private Sub RefreshGadgetManager(gadgetType, timeStr)
    Dim gadgetManager
    Set gadgetManager = GetGadgetManager(gadgetType)
    
    If gadgetManager Is Nothing Then
        Exit Sub
    End If
    
    Dim gadgets
    gadgets = gadgetManager.GetGadgets()
    
    If gadgets.Count > 0 Then
        Dim entries
        entries = gadgets.items
    
        Dim i
        For i = 0 To gadgets.Count - 1
            entries(i).RefreshData timeStr
        Next
    End If
End Sub

Public Function GetGadgetManager(gadgetType)
    If m_gadgetManagers.exists(gadgetType) Then
        Set GetGadgetManager = m_gadgetManagers.item(gadgetType)
    Else
        Set GetGadgetManager = Nothing
    End If
End Function

Public Function GetGadget(gadgetType, gadgetId)
    Set GetGadget = Nothing ' return nothing by default to indicate error
    
    Dim gadgetManager
    Set gadgetManager = GetGadgetManager(gadgetType)
    
    If gadgetManager Is Nothing Then
        Exit Function
    End If
    
    Dim gadgets
    Set gadgets = gadgetManager.GetGadgets()
    
    If Not gadgets.exists(gadgetId) Then
        Exit Function
    End If
    
    Set GetGadget = gadgets.item(gadgetId)
End Function
