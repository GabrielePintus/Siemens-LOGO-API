Attribute VB_Name = "Test"
Option Private Module
Public Sub TestTest()
    Dim packer As VariablePacker
    
    ' Case (Stand alone)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 1
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,1"
    
    ' Case (Stand alone)
    Set packer = New VariablePacker
    packer.AddVariable 1, 7, 7
    
    AssertEqual packer.GetRequestStr(), ",1,0,56,2,1"
    
    ' Case (Forward Merge)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 2
    packer.AddVariable 1, 1, 1
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,2"
    
    ' Case (Different Range)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 2
    packer.AddVariable 2, 1, 1
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,2;,2,0,8,2,1"
    
    ' Case (Merge)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 2
    packer.AddVariable 1, 13, 14
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,14"
    
    ' Case (Cannot Merge)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 2
    packer.AddVariable 1, 14, 15
    
    MsgBox packer.GetRequestStr()
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,2;,1,0,112,2,2"
    
    ' Case (Successive Merge)
    Set packer = New VariablePacker
    packer.AddVariable 1, 1, 2
    packer.AddVariable 1, 13, 14
    packer.AddVariable 1, 25, 25
    
    AssertEqual packer.GetRequestStr(), ",1,0,8,2,25"
End Sub

Public Sub AssertEqual(str1, str2)
    If str1 <> str2 Then
        MsgBox "Error Result(" + str2 + ") While Expecting(" + str1 + ")"
    End If
End Sub

Public Sub TestVariableData()
    Dim data As VariableData
    Set data = New VariableData
    
    data.FillData 1, 0, "FFFFFFFFFF"
    
    MsgBox data.GetValue(1, 0, VAR_SIZE_BIT)
    MsgBox data.GetValue(1, 0, VAR_SIZE_BYTE)
    MsgBox data.GetValue(1, 0, VAR_SIZE_WORD)
    MsgBox data.GetValue(1, 8, VAR_SIZE_DWORD)
End Sub

Public Sub TestMisc()
    MsgBox CDbl("4294967295")
End Sub


