Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-XMLConverter"
    
    On Error Resume Next
    
    Dim XMLString As String
    Dim XMLObject As Object
    
    ' ============================================= '
    ' ParseXML
    ' ============================================= '
    
    
    
    ' ============================================= '
    ' ConvertToXML
    ' ============================================= '
    
    
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    
    
    
    InlineRunner.RunSuite Specs
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite Specs
End Sub

Public Function ToMatchParseError(Actual As Variant, Args As Variant) As Variant
    Dim Partial As String
    Dim Arrow As String
    Dim Message As String
    Dim Description As String
    
    If UBound(Args) < 2 Then
        ToMatchParseError = "Need to pass expected partial, arrow, and message"
    ElseIf Err.Number = 10101 Then
        Partial = Args(0)
        Arrow = Args(1)
        Message = Args(2)
        Description = "Error parsing XML:" & vbNewLine & Partial & vbNewLine & Arrow & vbNewLine & Message
        
        Dim Parts As Variant
        Parts = Split(Err.Description, vbNewLine)
        
        If Parts(1) <> Partial Then
            ToMatchParseError = "Expected " & Parts(1) & " to equal " & Partial
        ElseIf Parts(2) <> Arrow Then
            ToMatchParseError = "Expected " & Parts(2) & " to equal " & Arrow
        ElseIf Parts(3) <> Message Then
            ToMatchParseError = "Expected " & Parts(3) & " to equal " & Message
        ElseIf Err.Description <> Description Then
            ToMatchParseError = "Expected " & Err.Description & " to equal " & Description
        Else
            ToMatchParseError = True
        End If
    Else
        ToMatchParseError = "Expected error number " & Err.Number & " to be 10101"
    End If
End Function
