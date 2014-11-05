Attribute VB_Name = "XMLConverter"
''
' VBA-XMLConverter v0.0.0
' (c) Tim Hall - https://github.com/timhall/VBA-JSONConverter
'
' XML Converter for VBA
'
' Errors:
' 10101 - XML parse error
'
' @author: tim.hall.engr@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
'
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

#If Mac Then
#ElseIf Win64 Then
Private Declare PtrSafe Sub XML_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (XML_MemoryDestination As Any, XML_MemorySource As Any, ByVal XML_ByteLength As Long)
#Else
Private Declare Sub XML_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (XML_MemoryDestination As Any, XML_MemorySource As Any, ByVal XML_ByteLength As Long)
#End If

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert XML string to object (Dictionary/Collection)
'
' @param {String} XML_String
' @return {Object} (Dictionary or Collection)
' -------------------------------------- '
Public Function ParseXML(ByVal XML_String As String, Optional XML_ConvertLargeNumbersToString As Boolean = True) As Object
    Dim XML_Index As Long
    XML_Index = 1
    
    ' Remove vbCr, vbLf, and vbTab from XML_String
    XML_String = VBA.Replace(VBA.Replace(VBA.Replace(XML_String, VBA.vbCr, ""), VBA.vbLf, ""), VBA.vbTab, "")
    
    XML_SkipSpaces XML_String, XML_Index
    Select Case VBA.Mid$(XML_String, XML_Index, 1)
    ' TODO
    Case Else
        ' Error: Invalid JSON string
        Err.Raise 10101, "XMLConverter", XML_ParseErrorMessage(XML_String, XML_Index, "...")
    End Select
End Function

''
' Convert object (Dictionary/Collection/Array) to XML
'
' @param {Variant} XML_DictionaryCollectionOrArray (Dictionary, Collection, or Array)
' @return {String}
' -------------------------------------- '
Public Function ConvertToJSON(ByVal XML_DictionaryCollectionOrArray As Variant, Optional XML_ConvertLargeNumbersFromString As Boolean = True) As String
    Dim XML_Buffer As String
    Dim XML_BufferPosition As Long
    Dim XML_BufferLength As Long
    
    ' TODO
End Function

' ============================================= '
' Private Functions
' ============================================= '



Private Function XML_Peek(XML_String As String, ByVal XML_Index As Long, Optional XML_NumberOfCharacters As Long = 1) As String
    ' "Peek" at the next number of characters without incrementing XML_Index (ByVal instead of ByRef)
    XML_SkipSpaces XML_String, XML_Index
    XML_Peek = VBA.Mid$(XML_String, XML_Index, XML_NumberOfCharacters)
End Function

Private Sub XML_SkipSpaces(XML_String As String, ByRef XML_Index As Long)
    ' Increment index to skip over spaces
    Do While XML_Index > 0 And XML_Index <= VBA.Len(XML_String) And VBA.Mid$(XML_String, XML_Index, 1) = " "
        XML_Index = XML_Index + 1
    Loop
End Sub

Private Function XML_StringIsLargeNumber(XML_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See XML_ParseNumber)
    
    Dim XML_Length As Long
    XML_Length = VBA.Len(XML_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If XML_Length >= 16 And XML_Length <= 100 Then
        Dim XML_CharCode As String
        Dim XML_Index As Long
        
        XML_StringIsLargeNumber = True
        
        For i = 1 To XML_Length
            XML_CharCode = VBA.Asc(VBA.Mid$(XML_String, i, 1))
            Select Case XML_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                XML_StringIsLargeNumber = False
                Exit Function
            End Select
        Next i
    End If
End Function

Private Function XML_ParseErrorMessage(XML_String As String, ByRef XML_Index As Long, ErrorMessage As String)
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing XML:
    ' TODO
    '          ^
    ' Expecting TODO
    
    Dim XML_StartIndex As Long
    Dim XML_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    XML_StartIndex = XML_Index - 10
    XML_StopIndex = XML_Index + 10
    If XML_StartIndex <= 0 Then
        XML_StartIndex = 1
    End If
    If XML_StopIndex > VBA.Len(XML_String) Then
        XML_StopIndex = VBA.Len(XML_String)
    End If

    XML_ParseErrorMessage = "Error parsing XML:" & VBA.vbNewLine & _
                             VBA.Mid$(XML_String, XML_StartIndex, XML_StopIndex - XML_StartIndex + 1) & VBA.vbNewLine & _
                             VBA.Space$(XML_Index - XML_StartIndex) & "^" & VBA.vbNewLine & _
                             ErrorMessage
End Function

Private Sub XML_BufferAppend(ByRef XML_Buffer As String, _
                              ByRef XML_Append As Variant, _
                              ByRef XML_BufferPosition As Long, _
                              ByRef XML_BufferLength As Long)
#If Mac Then
    XML_Buffer = XML_Buffer & XML_Append
#Else
    ' VBA can be slow to append strings due to allocating a new string for each append
    ' Instead of using the traditional append, allocate a large empty string and then copy string at append position
    '
    ' Example:
    ' Buffer: "abc  "
    ' Append: "def"
    ' Buffer Position: 3
    ' Buffer Length: 5
    '
    ' Buffer position + Append length > Buffer length -> Append chunk of blank space to buffer
    ' Buffer: "abc       "
    ' Buffer Length: 10
    '
    ' Copy memory for "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp

    Dim XML_AppendLength As Long
    Dim XML_LengthPlusPosition As Long
    
    XML_AppendLength = VBA.LenB(XML_Append)
    XML_LengthPlusPosition = XML_AppendLength + XML_BufferPosition
    
    If XML_LengthPlusPosition > XML_BufferLength Then
        ' Appending would overflow buffer, add chunks until buffer is long enough
        Dim XML_TemporaryLength As Long
        
        XML_TemporaryLength = XML_BufferLength
        Do While XML_TemporaryLength < XML_LengthPlusPosition
            ' Initially, initialize string with 255 characters,
            ' then add large chunks (8192) after that
            '
            ' Size: # Characters x 2 bytes / character
            If XML_TemporaryLength = 0 Then
                XML_TemporaryLength = XML_TemporaryLength + 510
            Else
                XML_TemporaryLength = XML_TemporaryLength + 16384
            End If
        Loop
        
        XML_Buffer = XML_Buffer & VBA.Space$((XML_TemporaryLength - XML_BufferLength) \ 2)
        XML_BufferLength = XML_TemporaryLength
    End If
    
    ' Copy memory from append to buffer at buffer position
    XML_CopyMemory ByVal XML_UnsignedAdd(StrPtr(XML_Buffer), _
                    XML_BufferPosition), _
                    ByVal StrPtr(XML_Append), _
                    XML_AppendLength
    
    XML_BufferPosition = XML_BufferPosition + XML_AppendLength
#End If
End Sub

Private Function XML_BufferToString(ByRef XML_Buffer As String, ByVal XML_BufferPosition As Long, ByVal XML_BufferLength As Long) As String
#If Mac Then
    XML_BufferToString = XML_Buffer
#Else
    If XML_BufferPosition > 0 Then
        XML_BufferToString = VBA.Left$(XML_Buffer, XML_BufferPosition \ 2)
    End If
#End If
End Function

#If Win64 Then
Private Function XML_UnsignedAdd(XML_Start As LongPtr, XML_Increment As Long) As LongPtr
#Else
Private Function XML_UnsignedAdd(XML_Start As Long, XML_Increment As Long) As Long
#End If

    If XML_Start And &H80000000 Then
        XML_UnsignedAdd = XML_Start + XML_Increment
    ElseIf (XML_Start Or &H80000000) < -XML_Increment Then
        XML_UnsignedAdd = XML_Start + XML_Increment
    Else
        XML_UnsignedAdd = (XML_Start + &H80000000) + (XML_Increment + &H80000000)
    End If
End Function
