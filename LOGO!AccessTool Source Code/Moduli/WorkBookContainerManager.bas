Attribute VB_Name = "WorkbookContainerManager"
Option Private Module
Dim m_WorkBookContainers

Public Function GetWorkBookContainers()
    'MsgBox "Called"
    
    If TypeName(m_WorkBookContainers) = "Empty" Then
        Set m_WorkBookContainers = CreateObject("Scripting.Dictionary")
        
        ' HeartBeat will start everything including the environment. This also ensure
        ' real time update of Hierarchy
        GetHeartBeatHandlers
        
        ' Initial Update will fill workbook containers
        UpdateHierarchy
    Else
        ' It seems the frequency of calling this function is not too high (From Top to Down for each operation)
        ' It is suggested to perform minimum safe check operation (Remove Obsolete Workbook Container).
        Dim keys
        keys = m_WorkBookContainers.keys
        
        Dim i
        For i = 0 To m_WorkBookContainers.Count - 1
            ' Remove those closed workbooks
            If TypeName(keys(i)) <> "Workbook" Then
                ' remove containers from the set
                m_WorkBookContainers.Remove keys(i)
            End If
            
            'MsgBox CStr(i) + ":" + TypeName(keys(i)) + ":" + TypeName(entries(i))
        Next
    End If
    
    Set GetWorkBookContainers = m_WorkBookContainers
End Function

Public Function GetWorkBookContainer(Workbook As Workbook)
    Dim containers
    Set containers = GetWorkBookContainers()
    
    If containers.exists(Workbook) Then
        Set GetWorkBookContainer = containers.item(Workbook)
    Else
        Set GetWorkBookContainer = Nothing
    End If
End Function

Public Function GetActiveWorkBookContainer()
    Set GetActiveWorkBookContainer = GetWorkBookContainer(Application.ActiveWorkbook)
End Function

Public Sub ActiveWorkBookEvent(gadgetType, gadgetId, eventId)
    Dim container
    Set container = GetActiveWorkBookContainer()
    If container Is Nothing Then
        Exit Sub
    End If
    
    Dim gadgetManager
    Set gadgetManager = container.GetGadgetManager(gadgetType)
    
    If gadgetManager Is Nothing Then
        Exit Sub
    End If
    
    gadgetManager.HandleEvent gadgetId, eventId
End Sub

Public Sub UpdateHierarchy()
    Dim i
    Dim keys
    Dim entries
    Dim containers
    
    Set containers = GetWorkBookContainers()

    entries = containers.items
    keys = containers.keys
    
    ' Update existing workbook containers
    For i = 0 To containers.Count - 1
        If Not entries(i).Update Then
            containers.Remove keys(i) ' Recycle the document container
        End If
    Next
    
    ' If all mapped, finish operation
    If containers.Count = Application.Workbooks.Count Then
        Exit Sub
    End If
    
    ' If not matching, add those not mapped ones
    For i = 1 To Application.Workbooks.Count
        Dim Workbook As Workbook
        Set Workbook = Application.Workbooks.item(i)
        
        If Not m_WorkBookContainers.exists(Workbook) Then
            Dim container As workbookContainer
            Set container = New workbookContainer
            container.Load Workbook ' Load will auto Update
            m_WorkBookContainers.Add Workbook, container
        End If
    Next
End Sub
