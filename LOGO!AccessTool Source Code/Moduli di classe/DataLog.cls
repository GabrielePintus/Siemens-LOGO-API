VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim m_logFile ' The handle to operate file
Dim m_count As Integer ' counter for each file
Dim m_dirPath As String
Dim m_prefix As String
Dim m_head As String
Const DATALOG_MAX_EACH_FILE As Integer = 20000
Const DATALOG_MAX_FILE_COUNT As Integer = 500

Public Function OpenDir(path As String)
    ' If new Path is the same as previous one, approve it simply
    ' current process will not be interruptted
    If path = m_dirPath Then
        OpenDir = True
        Exit Function
    End If

    Dim fso
    Set fso = GetFileSystemObject()
    
    ' close current file and reset dirPath at first
    CloseFile
    m_dirPath = ""
    
    ' update dirPath if the specified one is a valid one
    If fso.FolderExists(path) Then
        m_dirPath = path
        OpenDir = True
    Else
        OpenDir = False
    End If
End Function

Public Sub SetFilePrefix(prefix As String)
    If m_prefix <> prefix Then
        CloseFile
    
        m_prefix = prefix
    End If
End Sub

Public Sub SetHead(head As String)
    If m_head <> head Then
        CloseFile
    
        m_head = head
    End If
End Sub

Public Function IsSuccess()
    Dim file
    Set file = GetFile()
    
    If Not file Is Nothing Then
        IsSuccess = True
    Else
        IsSuccess = False
    End If
End Function

Public Sub WriteLog(data As String)
    Dim file
    Set file = GetFile()
    
    If Not file Is Nothing Then
        file.WriteLine (data)
        m_count = m_count + 1
    End If
End Sub

Public Sub CloseDir()
    CloseFile
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private implementations

Private Sub CloseFile()
    If IsFileOpened() Then
        m_logFile.Close
        Set m_logFile = Nothing
        m_count = 0
    End If
End Sub

Private Function IsFileOpened()
    If TypeName(m_logFile) = "TextStream" Then
        IsFileOpened = True
    Else
        IsFileOpened = False
    End If
End Function

Private Function GetFile()
    Set GetFile = Nothing ' return Nothing by default to indicate error
    
    Dim needCreateNew As Boolean

    ' The simplest operation, keep using it if it is still valid
    If IsFileOpened() Then
        If m_count < DATALOG_MAX_EACH_FILE Then
            Set GetFile = m_logFile
            Exit Function
        Else
            CloseFile
            needCreateNew = True
        End If
    End If
    
    ' -- comes here means we may need to create a new file
    
    ' Validate dirpath and head at first
    If m_dirPath = "" Then
        Exit Function
    End If
    
    If m_head = "" Then
        Exit Function
    End If
    
    ' do statistics about the folder before create / recycle / append determination
    Dim fso
    Set fso = GetFileSystemObject()
    
    Dim folder
    Dim fc
    Dim file
    
    Set folder = fso.GetFolder(m_dirPath)
    Set fc = folder.Files
    
    Dim eldestFile As String
    Dim latestFile As String
    Dim fileCount As Integer
    
    For Each file In fc
        Dim fileName
        fileName = file.name
        
        If Left(fileName, Len(m_prefix)) = m_prefix Then
            If IsDataLogFileName(Mid(fileName, Len(m_prefix) + 1)) Then
                ' Upadte latestFile and eldestFile record
                If fileCount = 0 Then
                    latestFile = fileName
                    eldestFile = fileName
                Else
                    If fileName > latestFile Then
                        latestFile = fileName
                    End If
                    
                    If fileName < eldestFile Then
                        eldestFile = fileName
                    End If
                End If
                
                ' Increment the file count
                fileCount = fileCount + 1
            End If
        End If
    Next
    
    ' Try to append if not necessary to create new one
    If (Not needCreateNew) And (fileCount > 0) Then
        If TryToAppendDataLog(fso, m_dirPath + "\" + latestFile) Then
            Set GetFile = m_logFile
            Exit Function
        End If
    End If
    
    ' -- comes here means we definitely need to create a new file
    
    ' Try to recycle the eldest one. No matter what happens, we will ingore errors
    If fileCount >= DATALOG_MAX_FILE_COUNT Then
        TryToRecycleDataLog fso, m_dirPath + "\" + eldestFile
    End If

    ' Try to create a new data log file
    Dim newFileName
    
    newFileName = m_dirPath + "\" + m_prefix + CStr(Format(Now, "yyyymmddhhmmss")) + ".csv"
    TryToOpenNewDataLog fso, newFileName

    ' use current logFile as return value
    Set GetFile = m_logFile
End Function

Private Function TryToAppendDataLog(fso, latestDataLogFileFullPath)
    On Error GoTo Err ' catch error in operation (e.g. the file is exclusively hold by another application)
    
    TryToAppendDataLog = False ' return false by default to indicate error

    If latestDataLogFileFullPath <> "" Then
        Dim fileObj
        Set fileObj = fso.OpenTextFile(latestDataLogFileFullPath, 1, False) ' For Reading, Not Create
        
        ' Get the head if it has some
        Dim head As String
        If Not fileObj.AtEndOfStream Then
            head = fileObj.ReadLine
        End If
        
        fileObj.Close
        
        ' If head doesn't match, don't take it
        If head <> m_head Then
            Exit Function
        End If
        
        Set fileObj = fso.OpenTextFile(latestDataLogFileFullPath, 8, False) ' For Appending, Not Create
        
        Dim dataCount
        
        dataCount = fileObj.Line - 2 ' remove index start & head
        
        If dataCount >= DATALOG_MAX_EACH_FILE Then
            fileObj.Close
            Exit Function
        End If
        
        ' Now we can say it is a valid file for appending
        Set m_logFile = fileObj
        m_count = dataCount
        TryToAppendDataLog = True ' return true to indicate successful operation
    End If
    
    Exit Function
    
Err:

End Function

Private Sub TryToRecycleDataLog(fso, eldestDataLogFileFullPath)
    On Error GoTo Err ' catch error in operation (e.g. the file is exclusively hold by another application)

    If eldestDataLogFileFullPath <> "" Then
        fso.DeleteFile eldestDataLogFileFullPath
    End If
    
    Exit Sub
    
Err:

End Sub

Private Sub TryToOpenNewDataLog(fso, newDataLogFileFullPath)
    On Error GoTo Err ' catch error in operation (e.g. the file is exclusively hold by another application)

    If newDataLogFileFullPath <> "" Then
        Set m_logFile = fso.OpenTextFile(newDataLogFileFullPath, 8, True)
        m_count = 0
        
        m_logFile.WriteLine (m_head)
    End If
    
    Exit Sub
    
Err:

End Sub



