Attribute VB_Name = "XMLConverter"
''
' VBA-XML v0.1.0
' (c) Tim Hall - https://github.com/VBA-tools/VBA-XML
'
' XML Converter for VBA
'
' Design:
' The goal is to have the general form of MSXML2.DOMDocument (albeit not feature complete)
'
' ParseXML(<messages><message id="1">A</message><message id="2">B</message></messages>) ->
'
' {Dictionary}
' - nodeName: {String} "#document"
' - attributes: {Collection} (Nothing)
' - childNodes: {Collection}
'   {Dictionary}
'   - nodeName: "messages"
'   - attributes: (empty)
'   - childNodes:
'     {Dictionary}
'     - nodeName: "message"
'     - attributes:
'       {Collection of Dictionary}
'       nodeName: "id"
'       text: "1"
'     - childNodes: (empty)
'     - text: A
'     {Dictionary}
'     - nodeName: "message"
'     - attributes:
'       {Collection of Dictionary}
'       nodeName: "id"
'       text: "2"
'     - childNodes: (empty)
'     - text: B
'
' Errors:
' 10101 - XML parse error
'
' References:
' - http://www.w3.org/TR/REC-xml/
'
' @author: tim.hall.engr@gmail.com
' @author: Andrew Pullon | andrew.pullon@pkfh.co.nz | andrewcpullon@gmail.com
' @license: MIT (http://www.opensource.org/licenses/mit-license.php
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '
Option Explicit

' === VBA-UTC Headers
#If Mac Then

#If VBA7 Then

' 64-bit Mac (2016)
Private Declare PtrSafe Function utc_popen Lib "/usr/lib/libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As LongPtr
Private Declare PtrSafe Function utc_pclose Lib "/usr/lib/libc.dylib" Alias "pclose" _
    (ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_fread Lib "/usr/lib/libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As LongPtr, ByVal utc_Number As LongPtr, ByVal utc_File As LongPtr) As LongPtr
Private Declare PtrSafe Function utc_feof Lib "/usr/lib/libc.dylib" Alias "feof" _
    (ByVal utc_File As LongPtr) As LongPtr

#Else

' 32-bit Mac
Private Declare Function utc_popen Lib "libc.dylib" Alias "popen" _
    (ByVal utc_Command As String, ByVal utc_Mode As String) As Long
Private Declare Function utc_pclose Lib "libc.dylib" Alias "pclose" _
    (ByVal utc_File As Long) As Long
Private Declare Function utc_fread Lib "libc.dylib" Alias "fread" _
    (ByVal utc_Buffer As String, ByVal utc_Size As Long, ByVal utc_Number As Long, ByVal utc_File As Long) As Long
Private Declare Function utc_feof Lib "libc.dylib" Alias "feof" _
    (ByVal utc_File As Long) As Long

#End If

#ElseIf VBA7 Then

' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724421.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms724949.aspx
' http://msdn.microsoft.com/en-us/library/windows/desktop/ms725485.aspx
Private Declare PtrSafe Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare PtrSafe Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare PtrSafe Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#Else

Private Declare Function utc_GetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION) As Long
Private Declare Function utc_SystemTimeToTzSpecificLocalTime Lib "kernel32" Alias "SystemTimeToTzSpecificLocalTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpUniversalTime As utc_SYSTEMTIME, utc_lpLocalTime As utc_SYSTEMTIME) As Long
Private Declare Function utc_TzSpecificLocalTimeToSystemTime Lib "kernel32" Alias "TzSpecificLocalTimeToSystemTime" _
    (utc_lpTimeZoneInformation As utc_TIME_ZONE_INFORMATION, utc_lpLocalTime As utc_SYSTEMTIME, utc_lpUniversalTime As utc_SYSTEMTIME) As Long

#End If

#If Mac Then

#If VBA7 Then
Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As LongPtr
End Type

#Else

Private Type utc_ShellResult
    utc_Output As String
    utc_ExitCode As Long
End Type

#End If

#Else

Private Type utc_SYSTEMTIME
    utc_wYear As Integer
    utc_wMonth As Integer
    utc_wDayOfWeek As Integer
    utc_wDay As Integer
    utc_wHour As Integer
    utc_wMinute As Integer
    utc_wSecond As Integer
    utc_wMilliseconds As Integer
End Type

Private Type utc_TIME_ZONE_INFORMATION
    utc_Bias As Long
    utc_StandardName(0 To 31) As Integer
    utc_StandardDate As utc_SYSTEMTIME
    utc_StandardBias As Long
    utc_DaylightName(0 To 31) As Integer
    utc_DaylightDate As utc_SYSTEMTIME
    utc_DaylightBias As Long
End Type

#End If
' === End VBA-UTC

Private Type xml_Options
    ' VBA only stores 15 significant digits, so any numbers larger than that are truncated
    ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
    ' See: http://support.microsoft.com/kb/269370
    '
    ' By default, VBA-XML will use String for numbers longer than 15 characters that contain only digits
    ' to override set this option to `True`.
    UseDoubleForLargeNumbers As Boolean
    
    ' Use this option to include Node mapping (`parentNode`, `firstChild`, `lastChild`) in parsed object.
    ' Performance suffers (slightly) when including node mapping in object structure.
    IncludeNodeMapping As Boolean
    
    ' Internal VBA-XML parser is much slower than using `MSXML2.DOMDocument`. By default on Windows
    ' machines `MSXML2.DOMDocument` is used. Set this option to `True` to force use of VBA-XML.
    ' Not recommended if dealing with large XML strings (>1,000,000 char).
    '
    ' This option has no effect on Mac machines.
    ForceVbaXml As Boolean
End Type
Public XmlOptions As xml_Options

' ============================================= '
' Public Methods
' ============================================= '

''
' Convert XML string to Dictionary or DOMDocument (windows only).
'
' @method ParseXml
' @param {String} XmlString
' @return {DOMDocument|Dictionary}
''
Public Function ParseXml(ByVal XmlString As String) As Object
    Dim xml_String As String
    Dim xml_Index As Long
    xml_Index = 1

    ' Remove vbCr, vbLf, and vbTab from xml_String
    xml_String = VBA.Replace(VBA.Replace(VBA.Replace(XmlString, VBA.vbCr, vbNullString), VBA.vbLf, vbNullString), VBA.vbTab, vbNullString)

    xml_SkipSpaces xml_String, xml_Index
    If Not VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
        ' Error: Invalid XML string
        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
    Else
#If Mac Then
        Set ParseXml = New Dictionary
        ParseXml.Add "prolog", xml_ParseProlog(xml_String, xml_Index)
        ParseXml.Add "doctype", xml_ParseDoctype(xml_String, xml_Index)
        ParseXml.Add "nodeName", "#document"
        ParseXml.Add "attributes", Nothing
        ParseXml.Add "childNodes", New Collection
        ParseXml.Item("childNodes").Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, ParseXml, Nothing))
#Else
        If XmlOptions.ForceVbaXml Then
            Set ParseXml = New Dictionary
            ParseXml.Add "prolog", xml_ParseProlog(xml_String, xml_Index)
            ParseXml.Add "doctype", xml_ParseDoctype(xml_String, xml_Index)
            ParseXml.Add "nodeName", "#document"
            ParseXml.Add "attributes", Nothing
            ParseXml.Add "childNodes", New Collection
            ParseXml.Item("childNodes").Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, ParseXml, Nothing))
        Else
            Set ParseXml = CreateObject("MSXML2.DOMDocument")
            ParseXml.Async = False
            ParseXml.LoadXML XmlString
        End If
#End If
    End If
End Function

''
' Convert object (Dictionary/Collection/DOMDocument) to XML string.
'
' @method ConvertToXml
' @param {Variant} XmlValue (Dictionary, Collection, or DOMDocument)
' @param {Integer|String} Whitespace "Pretty" print xmlwith given number of spaces per indentation (Integer) or given string
' @return {String}
''
Public Function ConvertToXML(ByVal XmlValue As Variant, Optional ByVal Whitespace As Variant, Optional ByVal xml_CurrentIndentation As Long = 0) As String
    Dim xml_buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    Dim xml_Indentation As String
    Dim xml_PrettyPrint As Boolean
    Dim xml_Converted As String
    Dim xml_ChildNode As Dictionary
    Dim xml_Attribute As Dictionary
    
    xml_PrettyPrint = Not IsMissing(Whitespace)
    
    Select Case VBA.VarType(XmlValue)
    Case VBA.vbNull
        ConvertToXML = vbNullString
    Case VBA.vbDate
        ConvertToXML = ConvertToIso(VBA.CDate(XmlValue))
    Case VBA.vbString
        If Not XmlOptions.UseDoubleForLargeNumbers And xml_StringIsLargeNumber(XmlValue) Then
            ConvertToXML = XmlValue
        Else
            ConvertToXML = xml_Encode(XmlValue)
        End If
    Case VBA.vbBoolean
        ConvertToXML = VBA.IIf(XmlValue, "true", "false")
    Case VBA.vbObject
        If xml_PrettyPrint Then
            If VBA.VarType(Whitespace) = VBA.vbString Then
                xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
            Else
                xml_Indentation = VBA.Space$((xml_CurrentIndentation) * Whitespace)
            End If
        End If
        
        ' Dictionary (Node).
        If VBA.TypeName(XmlValue) = "Dictionary" Then
            ' If root node, parse prolog and child nodes then exit.
            If XmlValue.Item("nodeName") = "#document" Then
                If Not XmlValue.Item("prolog") = vbNullString Then
                    xml_BufferAppend xml_buffer, XmlValue.Item("prolog"), xml_BufferPosition, xml_BufferLength
                    If xml_PrettyPrint Then
                        xml_BufferAppend xml_buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                    End If
                End If
                xml_Converted = ConvertToXML(XmlValue.Item("childNodes"), Whitespace, xml_CurrentIndentation)
                xml_BufferAppend xml_buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                ConvertToXML = xml_BufferToString(xml_buffer, xml_BufferPosition)
                Exit Function
            Else
                ' Validate Dictionary structure.
                If Not XmlValue.Exists("nodeName") Or Not XmlValue.Exists("nodeValue") Then
                    Err.Raise 11001, "XMLConverter", "Error parsing XML:" & VBA.vbNewLine & Err.Number & " - " & Err.Description & _
                        "Poorly structured XML Dictionary. Use `ParseXml` with `XmlOptions.ForceVbaXml = True` OR " & _
                        "`CreateNode` and `CreateAttribute` to create a correctly structured XML dictionary object."
                End If
            
                ' Add 'Start Tag'.
                xml_BufferAppend xml_buffer, xml_Indentation & "<", xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_buffer, XmlValue.Item("nodeName"), xml_BufferPosition, xml_BufferLength
                If XmlValue.Exists("attributes") Then
                    If Not XmlValue.Item("attributes") Is Nothing Then
                        For Each xml_Attribute In XmlValue.Item("attributes")
                            xml_BufferAppend xml_buffer, " " & xml_Attribute.Item("name") & "=" & """" & xml_Attribute.Item("value") & """", xml_BufferPosition, xml_BufferLength
                        Next xml_Attribute
                    End If
                End If
                
                ' Check for void node.
                If xml_IsVoidNode(XmlValue) Then
                    ' Add 'Empty Element' tag and exit.
                    xml_BufferAppend xml_buffer, "/>", xml_BufferPosition, xml_BufferLength
                    If xml_PrettyPrint Then
                        xml_BufferAppend xml_buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                        
                        If VBA.VarType(Whitespace) = VBA.vbString Then
                            xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                        Else
                            xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                        End If
                    End If
                    ConvertToXML = xml_BufferToString(xml_buffer, xml_BufferPosition)
                    Exit Function
                Else
                    ' Finish 'Start Tag' and continue.
                    xml_BufferAppend xml_buffer, ">", xml_BufferPosition, xml_BufferLength
                End If
                
                ' Add node content.
                If XmlValue.Exists("childNodes") Then
                    If XmlValue.Item("childNodes").Count > 0 Then
                        If xml_PrettyPrint Then
                            xml_BufferAppend xml_buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
            
                            If VBA.VarType(Whitespace) = VBA.vbString Then
                                xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                            Else
                                xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                            End If
                        End If
                    
                        ' Convert childNodes.
                        xml_Converted = ConvertToXML(XmlValue.Item("childNodes"), Whitespace, xml_CurrentIndentation + 1)
                        xml_BufferAppend xml_buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                        xml_BufferAppend xml_buffer, xml_Indentation, xml_BufferPosition, xml_BufferLength
                    Else
                        ' No child nodes, add text.
                        xml_Converted = ConvertToXML(XmlValue.Item("nodeValue"), Whitespace, xml_CurrentIndentation + 1)
                        xml_BufferAppend xml_buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                    End If
                Else
                    ' No child nodes, add text.
                    xml_Converted = ConvertToXML(XmlValue.Item("nodeValue"), Whitespace, xml_CurrentIndentation + 1)
                    xml_BufferAppend xml_buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                End If
                
                ' Add 'End Tag'.
                xml_BufferAppend xml_buffer, "</", xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_buffer, XmlValue.Item("nodeName"), xml_BufferPosition, xml_BufferLength
                xml_BufferAppend xml_buffer, ">", xml_BufferPosition, xml_BufferLength
                
                If xml_PrettyPrint Then
                    xml_BufferAppend xml_buffer, vbNewLine, xml_BufferPosition, xml_BufferLength
                    
                    If VBA.VarType(Whitespace) = VBA.vbString Then
                        xml_Indentation = VBA.String$(xml_CurrentIndentation, Whitespace)
                    Else
                        xml_Indentation = VBA.Space$(xml_CurrentIndentation * Whitespace)
                    End If
                End If
            End If
            ConvertToXML = xml_BufferToString(xml_buffer, xml_BufferPosition)
        
        ' Collection (child nodes)
        ElseIf VBA.TypeName(XmlValue) = "Collection" Then
            For Each xml_ChildNode In XmlValue
                ' Convert node.
                xml_Converted = ConvertToXML(xml_ChildNode, Whitespace, xml_CurrentIndentation)
                If Not xml_Converted = vbNullString Then
                    xml_BufferAppend xml_buffer, xml_Converted, xml_BufferPosition, xml_BufferLength
                Else
                    xml_BufferAppend xml_buffer, "null", xml_BufferPosition, xml_BufferLength
                End If
            Next xml_ChildNode
            
            ConvertToXML = xml_BufferToString(xml_buffer, xml_BufferPosition)
        
        ' MSXML2.DOMDocument (windows only)
        ElseIf VBA.TypeName(XmlValue) = "DOMDocument" Then
            If xml_PrettyPrint Then
                Dim xml_Writer As Object
                Dim xml_Reader As Object
                
                Set xml_Writer = CreateObject("MSXML2.MXXMLWriter")
                Set xml_Reader = CreateObject("MSXML2.SAXXMLReader")
                
                xml_Writer.Indent = True
                Set xml_Reader.contentHandler = xml_Writer
                xml_Reader.Parse XmlValue.XML
                
                ConvertToXML = xml_Writer.Output
            Else
                ConvertToXML = VBA.Trim$(VBA.Replace(VBA.Replace(XmlValue.XML, vbCrLf, vbNullString), vbTab, vbNullString))
            End If
        End If
    Case VBA.vbInteger, VBA.vbLong, VBA.vbSingle, VBA.vbDouble, VBA.vbCurrency, VBA.vbDecimal
        ' Number (use decimals for numbers)
        ConvertToXML = VBA.Replace(XmlValue, ",", ".")
    Case Else
        ' vbEmpty, vbError, vbDataObject, vbByte, vbUserDefinedType
        ' Use VBA's built-in to-string
        On Error Resume Next
        ConvertToXML = XmlValue
        On Error GoTo 0
    End Select
    Exit Function
End Function

' ============================================= '
' Private Functions
' ============================================= '

Private Function xml_ParseProlog(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_OpeningLevel As Long
    Dim xml_StringLength As Long
    Dim xml_StartIndex As Long
    Dim xml_Chars As String

    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 2) = "<?" Then
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 2
        xml_StringLength = Len(xml_String)
    
        ' Find matching closing tag, ?>
        Do
            xml_Chars = VBA.Mid$(xml_String, xml_Index, 2)
            
            If xml_Index + 1 > xml_StringLength Then
                Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '?>'")
            ElseIf xml_OpeningLevel = 0 And xml_Chars = "?>" Then
                xml_Index = xml_Index + 2
                Exit Do
            ElseIf xml_Chars = "<?" Then
                xml_OpeningLevel = xml_OpeningLevel + 1
                xml_Index = xml_Index + 2
            ElseIf xml_Chars = "?>" Then
                xml_OpeningLevel = xml_OpeningLevel - 1
                xml_Index = xml_Index + 2
            Else
                xml_Index = xml_Index + 1
            End If
        Loop
        
        xml_ParseProlog = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

Private Function xml_ParseDoctype(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_OpeningLevel As Long
    Dim xml_StringLength As Long
    Dim xml_StartIndex As Long
    Dim xml_Char As String
    
    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 2) = "<!" Then
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 2
        xml_StringLength = Len(xml_String)
        
        ' Find matching closing tag, >
        Do
            xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
            xml_Index = xml_Index + 1
            
            If xml_Index > xml_StringLength Then
                Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '>'")
            ElseIf xml_OpeningLevel = 0 And xml_Char = ">" Then
                Exit Do
            ElseIf xml_Char = "<" Then
                xml_OpeningLevel = xml_OpeningLevel + 1
            ElseIf xml_Char = ">" Then
                xml_OpeningLevel = xml_OpeningLevel - 1
            End If
        Loop
        
        xml_ParseDoctype = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

''
' Parse Node Attributes.
'
' <title lang="en">Harry Potter</title>
'       ^         ^
'     Start      End
'
' {Dictionary} Attribute
' -> Key: Name  Value: lang
' -> Key: Value Value: en
'
' @method xml_ParseAttributes
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {Collection} Collection of attributes (Dictionary).
''
Private Function xml_ParseAttributes(xml_String As String, ByRef xml_Index As Long) As Collection
    Dim xml_Char As String
    Dim xml_StartIndex As Long
    Dim xml_Quote As String
    Dim xml_Name As String
    
    Set xml_ParseAttributes = New Collection
    xml_SkipSpaces xml_String, xml_Index
    xml_StartIndex = xml_Index
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
        
        Select Case xml_Char
        Case "="
            If xml_Name = vbNullString Then
                ' Found end of attribute name
                ' Extract name, skip '=', find quote char, reset start index
                xml_Name = VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
                xml_Index = xml_Index + 1
                xml_Quote = VBA.Mid$(xml_String, xml_Index, 1)
                xml_Index = xml_Index + 1
                xml_StartIndex = xml_Index
                
                ' Check for valid quote style of attribute value
                If Not xml_Quote = """" And Not xml_Quote = "'" Then
                    ' Invalid Attribute quote.
                    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ''' or '""'")
                End If
            Else
                ' '=' exists within attribute value. Continue.
                xml_Index = xml_Index + 1
            End If
        Case xml_Quote
            ' Found end of attribute value
            ' Store name, value as new attribute.
            With xml_ParseAttributes
                .Add New Dictionary
                .Item(.Count).Add "name", xml_Name
                .Item(.Count).Add "value", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
            End With
            
            ' Reset variables.
            xml_Name = vbNullString
            xml_Quote = vbNullString
            
            ' Increment.
            xml_Index = xml_Index + 1
            xml_SkipSpaces xml_String, xml_Index
            xml_StartIndex = xml_Index
            
            ' Check for end of tag.
            If VBA.Mid$(xml_String, xml_Index, 1) = ">" Or VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
                Exit Function ' End of tag, exit.
            End If
        Case Else
            xml_Index = xml_Index + 1
        End Select
    Loop
    
    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '>' or '/>'")
End Function

Private Function xml_ParseNode(xml_String As String, ByRef xml_Index As Long, Optional ByRef xml_Parent As Dictionary) As Dictionary
    Dim xml_StartIndex As Long
    
    xml_SkipSpaces xml_String, xml_Index
    If VBA.Mid$(xml_String, xml_Index, 1) <> "<" Then
        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
    Else
        ' Skip opening bracket
        xml_StartIndex = xml_Index
        xml_Index = xml_Index + 1
        
        ' Initialize node
        Set xml_ParseNode = New Dictionary
        If XmlOptions.IncludeNodeMapping Then
            xml_ParseNode.Add "parentNode", xml_Parent
        End If
        xml_ParseNode.Add "attributes", Nothing
        xml_ParseNode.Add "childNodes", New Collection
        xml_ParseNode.Add "text", vbNullString
        If XmlOptions.IncludeNodeMapping Then
            xml_ParseNode.Add "firstChild", Nothing
            xml_ParseNode.Add "lastChild", Nothing
        End If
        xml_ParseNode.Add "nodeValue", Null
        
        ' 1. Parse nodeName
        xml_ParseNode.Add "nodeName", xml_ParseName(xml_String, xml_Index)
        
        ' 2. Parse attributes
        If VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
            ' '/>' is the 'Empty-element' tag. Nothing more to parse. Skip over closing '/>' and exit
            xml_Index = xml_Index + 2
            xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex) ' Add 'xml' text.
            Exit Function
        ElseIf VBA.Mid$(xml_String, xml_Index, 1) = ">" Then
            ' If '>' then end of Start Tag. Skip over closing '>' and continue.
            xml_Index = xml_Index + 1
        Else
            ' If not '/>' or '>' then attributes are present within Start Tag.
            Set xml_ParseNode.Item("attributes") = xml_ParseAttributes(xml_String, xml_Index)
            
            ' Re-do previous checks as index has moved to after attributes.
            If VBA.Mid$(xml_String, xml_Index, 2) = "/>" Then
                ' '/>' is the 'Empty-element' tag. Nothing more to parse. Skip over closing '/>' and exit
                xml_Index = xml_Index + 2
                xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex) ' Add 'xml' text.
                Exit Function
            ElseIf VBA.Mid$(xml_String, xml_Index, 1) = ">" Then
                ' If '>' then end of Start Tag. Skip over closing '>' and continue.
                xml_Index = xml_Index + 1
            End If
        End If
    
        ' 3. Parse node content (child nodes, text, value).
        xml_SkipSpaces xml_String, xml_Index
        If Not VBA.Mid$(xml_String, xml_Index, 2) = "</" Then
            If VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
                ' If '<' (but not '</'), then child node exists.
                If XmlOptions.IncludeNodeMapping Then
                    Set xml_ParseNode.Item("childNodes") = xml_ParseChildNodes(xml_String, xml_Index, xml_ParseNode)
                    Set xml_ParseNode.Item("firstChild") = xml_ParseNode.Item("childNodes").Item(1)
                    Set xml_ParseNode.Item("lastChild") = xml_ParseNode.Item("childNodes").Item(xml_ParseNode.Item("childNodes").Count)
                Else
                    Set xml_ParseNode.Item("childNodes") = xml_ParseChildNodes(xml_String, xml_Index)
                End If
                ' Set node 'Text' once child nodes are parsed (node text is space separated text of all child nodes).
                Dim xml_buffer As String
                Dim xml_BufferPosition As Long
                Dim xml_BufferLength As Long
                Dim xml_ChildNode As Dictionary
                For Each xml_ChildNode In xml_ParseNode.Item("childNodes")
                    xml_BufferAppend xml_buffer, xml_ChildNode.Item("text"), xml_BufferPosition, xml_BufferLength
                    xml_BufferAppend xml_buffer, " ", xml_BufferPosition, xml_BufferLength
                Next xml_ChildNode
                xml_ParseNode.Item("text") = xml_BufferToString(xml_buffer, xml_BufferPosition - 1)
            Else
                ' No child nodes. Set Node Text.
                xml_ParseNode.Item("text") = xml_ParseText(xml_String, xml_Index)
                xml_ParseNode.Item("nodeValue") = xml_ParseValue(xml_ParseNode.Item("text"))
            End If
        End If

        ' Skip over End-Tag '</' + 'nodeName' + '>'.
        xml_Index = xml_Index + 2 + VBA.Len(xml_ParseNode.Item("nodeName")) + 1
        
        ' Add 'xml' text.
        xml_ParseNode.Add "xml", VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex)
    End If
End Function

''
' Call 'xml_ParseNode' to parse each child node.
'
'  <book category="cooking">
'    <title lang="en">Everyday Italian</title>
'    ^
'  Start
'    <author>Giada De Laurentiis</author>
'    <year>2005</year>
'    <price>30.00</price>
'  </book>
'  ^
' End
'
' @method xml_ParseChildNodes
' @param {Dictionary} xml_Node | Parent Node.
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index |  Current index position in XML string.
''
Private Function xml_ParseChildNodes(xml_String As String, ByRef xml_Index As Long, Optional ByRef xml_Parent As Dictionary) As Collection
    Set xml_ParseChildNodes = New Collection
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_SkipSpaces xml_String, xml_Index
        If VBA.Mid$(xml_String, xml_Index, 2) = "</" Then
            Exit Function
        ElseIf VBA.Mid$(xml_String, xml_Index, 1) = "<" Then
            xml_ParseChildNodes.Add xml_ParseNode(xml_String, xml_Index, VBA.IIf(XmlOptions.IncludeNodeMapping, xml_Parent, Nothing))
        Else
            Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '</' or '<'")
        End If
    Loop
End Function

''
' Parse Node Name.
'
' <title lang="en">Harry Potter</title>
'  ^     ^
'Start  End  |   Name --> 'title'
'
' <author>Giada De Laurentiis</author>
'  ^     ^
'Start  End  |   Name --> 'author'
'
' <price />
'  ^     ^
'Start  End  |   Name --> 'price'
'
' @method xml_ParseName
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {String} nodeName
''
Private Function xml_ParseName(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_Char As String
    Dim xml_buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    
    xml_SkipSpaces xml_String, xml_Index
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
        
        Select Case xml_Char
        Case " ", ">", "/"
            xml_ParseName = xml_BufferToString(xml_buffer, xml_BufferPosition)
            If xml_Char = " " Then xml_Index = xml_Index + 1 ' Skip space
            Exit Function
        Case Else
            xml_BufferAppend xml_buffer, xml_Char, xml_BufferPosition, xml_BufferLength
            xml_Index = xml_Index + 1
        End Select
    Loop
            
    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ' ', '>', or '/>'")
End Function

''
' Parse Node text.
'
' <title lang="en">Harry Potter</title>
'                  ^           ^
'                Start        End
' Text --> 'Harry Potter'
'
' @method xml_ParseText
' @param {String} xml_String | Complete XML string to parse.
' @param {Long} xml_Index | Current index position in XML string.
' @return {String} Node text
''
Private Function xml_ParseText(xml_String As String, ByRef xml_Index As Long) As String
    Dim xml_Char As String
    Dim xml_buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    Dim xml_StartIndex As Long
    Dim xml_EncodedFound As Boolean
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
        xml_Char = VBA.Mid$(xml_String, xml_Index, 1)

        Select Case xml_Char
        Case "<" 'Closing tag.
            xml_ParseText = xml_BufferToString(xml_buffer, xml_BufferPosition)
            Exit Function
        Case "&"
            ' Remove encoding from XML string. See `xml_Encode` for additional information.
            ' Store start of encoded char and continue.
            xml_StartIndex = xml_Index
            xml_Index = xml_Index + 1
            xml_EncodedFound = False
            ' Find close of encoded char.
            Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String)
                xml_Char = VBA.Mid$(xml_String, xml_Index, 1)
                Select Case xml_Char
                Case ";"
                    xml_EncodedFound = True
                    Select Case VBA.Mid$(xml_String, xml_StartIndex, xml_Index - xml_StartIndex + 1)
                    Case "&quot;"
                        xml_BufferAppend xml_buffer, """", xml_BufferPosition, xml_BufferLength
                    Case "&amp;"
                        xml_BufferAppend xml_buffer, "&", xml_BufferPosition, xml_BufferLength
                    Case "&apos;"
                        xml_BufferAppend xml_buffer, "'", xml_BufferPosition, xml_BufferLength
                    Case "&lt;"
                        xml_BufferAppend xml_buffer, "<", xml_BufferPosition, xml_BufferLength
                    Case "&gt;"
                        xml_BufferAppend xml_buffer, ">", xml_BufferPosition, xml_BufferLength
                    Case Else
                        Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '&quot;', '&amp;', '&apos;', '&lt;' or '&gt;'")
                    End Select
                    xml_Index = xml_Index + 1
                    Exit Do
                Case Else
                    xml_Index = xml_Index + 1
                End Select
            Loop
            If Not xml_EncodedFound Then Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting ';'")
        Case Else
            xml_BufferAppend xml_buffer, xml_Char, xml_BufferPosition, xml_BufferLength
            xml_Index = xml_Index + 1
        End Select
    Loop

    Err.Raise 10101, "XMLConverter", xml_ParseErrorMessage(xml_String, xml_Index, "Expecting '<'")
End Function

''
' Parse node 'text' to nodeValue. (i.e., String to Boolean, Double, Date).
'
' @method xml_ParseValue
' @param {String} xml_Text | Text to parse.
' @return {Variant} Node Value
''
Private Function xml_ParseValue(ByVal xml_Text As String) As Variant
    If xml_Text = "true" Then
        xml_ParseValue = True
    ElseIf xml_Text = "false" Then
        xml_ParseValue = False
    ElseIf xml_Text = "null" Then
        xml_ParseValue = Null
    ElseIf VBA.IsNumeric(xml_Text) Then
        xml_ParseValue = xml_ParseNumber(xml_Text)
    ElseIf VBA.IsNumeric(VBA.Replace(VBA.Left$(xml_Text, 10), "-", vbNullString)) And VBA.InStr(xml_Text, "T") And VBA.IIf(VBA.InStr(xml_Text, "Z"), VBA.Len(xml_Text) = 20, VBA.Len(xml_Text) = 19) Then
        xml_ParseValue = ParseIso(xml_Text)
    Else
        xml_ParseValue = xml_Text
    End If
End Function

Private Function xml_ParseNumber(ByVal xml_Text As String) As Variant
    Dim xml_Index As Long
    Dim xml_Char As String
    Dim xml_Value As String
    Dim xml_IsLargeNumber As Boolean
    Dim xml_IsGUID As Boolean
    Dim xml_IsISODate As Boolean
    
    xml_Index = 1
    
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_Text) + 1
        xml_Char = VBA.Mid$(xml_Text, xml_Index, 1)

        If VBA.InStr("+-0123456789.eE", xml_Char) And Not xml_Char = vbNullString Then
            ' Unlikely to have massive number, so use simple append rather than buffer here
            xml_Value = xml_Value & xml_Char
            xml_Index = xml_Index + 1
        Else
            ' Excel only stores 15 significant digits, so any numbers larger than that are truncated
            ' This can lead to issues when BIGINT's are used (e.g. for Ids or Credit Cards), as they will be invalid above 15 digits
            ' See: http://support.microsoft.com/kb/269370
            '
            ' Fix: Parse -> String, Convert -> String longer than 15/16 characters containing only numbers and decimal points -> Number
            ' (decimal doesn't factor into significant digit count, so if present check for 15 digits + decimal = 16)
            xml_IsLargeNumber = VBA.IIf(VBA.InStr(xml_Value, "."), VBA.Len(xml_Value) >= 17, VBA.Len(xml_Value) >= 16)
            If Not XmlOptions.UseDoubleForLargeNumbers And xml_IsLargeNumber Then
                xml_ParseNumber = xml_Value
            Else
                ' VBA.Val does not use regional settings, so guard for comma is not needed
                xml_ParseNumber = VBA.Val(xml_Value)
            End If
            Exit Function
        End If
    Loop
End Function

Private Function xml_IsVoidNode(ByVal xml_Node As Dictionary) As Boolean
    If xml_Node.Exists("childNodes") Then
        xml_IsVoidNode = VBA.IsNull(xml_Node.Item("nodeValue")) And xml_Node.Item("childNodes").Count = 0
    Else
        xml_IsVoidNode = VBA.IsNull(xml_Node.Item("nodeValue"))
    End If
End Function

Private Function xml_Encode(ByVal xml_Text As Variant) As String
    ' Variables.
    Dim xml_Index As Long
    Dim xml_Char As String
    Dim xml_AscCode As Long
    Dim xml_buffer As String
    Dim xml_BufferPosition As Long
    Dim xml_BufferLength As Long
    
    For xml_Index = 1 To VBA.Len(xml_Text)
        xml_Char = VBA.Mid$(xml_Text, xml_Index, 1)
        xml_AscCode = VBA.AscW(xml_Char)

        ' When AscW returns a negative number, it returns the twos complement form of that number.
        ' To convert the twos complement notation into normal binary notation, add 0xFFF to the return result.
        ' https://support.microsoft.com/en-us/kb/272138
        If xml_AscCode < 0 Then
            xml_AscCode = xml_AscCode + 65536
        End If

        ' From spec, <, >, &, ", ' characters must be modified.
        Select Case xml_AscCode
        Case 34
            ' " -> 34 -> &quot;
            xml_Char = "&quot;"
        Case 38
            ' & -> 38 -> &amp;
            xml_Char = "&amp;"
        Case 39
            ' ' -> 39 -> &apos;
            xml_Char = "&apos;"
        Case 60
            ' < -> 60 -> &lt;
            xml_Char = "&lt;"
        Case 62
            ' > -> 62 -> &gt;
            xml_Char = "&gt;"
        End Select
        
        xml_BufferAppend xml_buffer, xml_Char, xml_BufferPosition, xml_BufferLength
    Next xml_Index
    
    xml_Encode = xml_BufferToString(xml_buffer, xml_BufferPosition)
End Function

Private Sub xml_SkipSpaces(ByVal xml_String As String, ByRef xml_Index As Long)
    ' Increment index to skip over spaces
    Do While xml_Index > 0 And xml_Index <= VBA.Len(xml_String) And VBA.Mid$(xml_String, xml_Index, 1) = " "
        xml_Index = xml_Index + 1
    Loop
End Sub

Private Function xml_StringIsLargeNumber(ByVal xml_String As Variant) As Boolean
    ' Check if the given string is considered a "large number"
    ' (See xml_ParseNumber)
    
    Dim xml_Length As Long
    Dim xml_CharIndex As Long
    xml_Length = VBA.Len(xml_String)
    
    ' Length with be at least 16 characters and assume will be less than 100 characters
    If xml_Length >= 16 And xml_Length <= 100 Then
        Dim xml_CharCode As String
        
        xml_StringIsLargeNumber = True
        
        For xml_CharIndex = 1 To xml_Length
            xml_CharCode = VBA.Asc(VBA.Mid$(xml_String, xml_CharIndex, 1))
            Select Case xml_CharCode
            ' Look for .|0-9|E|e
            Case 46, 48 To 57, 69, 101
                ' Continue through characters
            Case Else
                xml_StringIsLargeNumber = False
                Exit Function
            End Select
        Next xml_CharIndex
    End If
End Function

Private Function xml_ParseErrorMessage(ByVal xml_String As String, ByVal xml_Index As Long, ByVal xml_ErrorMessage As String) As String
    ' Provide detailed parse error message, including details of where and what occurred
    '
    ' Example:
    ' Error parsing XML:
    ' <abc>1234</def>
    '          ^
    ' Expecting '</abc>'
    
    Dim xml_StartIndex As Long
    Dim xml_StopIndex As Long
    
    ' Include 10 characters before and after error (if possible)
    xml_StartIndex = xml_Index - 10
    xml_StopIndex = xml_Index + 10
    If xml_StartIndex <= 0 Then
        xml_StartIndex = 1
    End If
    If xml_StopIndex > VBA.Len(xml_String) Then
        xml_StopIndex = VBA.Len(xml_String)
    End If

    xml_ParseErrorMessage = "Error parsing XML:" & VBA.vbNewLine & _
        VBA.Mid$(xml_String, xml_StartIndex, xml_StopIndex - xml_StartIndex + 1) & VBA.vbNewLine & _
        VBA.Space$(xml_Index - xml_StartIndex) & "^" & VBA.vbNewLine & _
        xml_ErrorMessage
End Function

Private Sub xml_BufferAppend(ByRef xml_buffer As String, _
                             ByRef xml_Append As Variant, _
                             ByRef xml_BufferPosition As Long, _
                             ByRef xml_BufferLength As Long)
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
    ' Put "def" into buffer at position 3 (0-based)
    ' Buffer: "abcdef    "
    '
    ' Approach based on cStringBuilder from vbAccelerator
    ' http://www.vbaccelerator.com/home/VB/Code/Techniques/RunTime_Debug_Tracing/VB6_Tracer_Utility_zip_cStringBuilder_cls.asp
    '
    ' and clsStringAppend from Philip Swannell
    ' https://github.com/VBA-tools/VBA-JSON/pull/82

    Dim xml_AppendLength As Long
    Dim xml_LengthPlusPosition As Long

    If Not xml_Append = vbNullString Then
        xml_AppendLength = VBA.Len(xml_Append)
        xml_LengthPlusPosition = xml_AppendLength + xml_BufferPosition
    
        If xml_LengthPlusPosition > xml_BufferLength Then
            ' Appending would overflow buffer, add chunk
            ' (double buffer length or append length, whichever is bigger)
            Dim xml_AddedLength As Long
            xml_AddedLength = VBA.IIf(xml_AppendLength > xml_BufferLength, xml_AppendLength, xml_BufferLength)
    
            xml_buffer = xml_buffer & VBA.Space$(xml_AddedLength)
            xml_BufferLength = xml_BufferLength + xml_AddedLength
        End If
    
        ' Note: Namespacing with VBA.Mid$ doesn't work properly here, throwing compile error:
        ' Function call on left-hand side of assignment must return Variant or Object
        Mid$(xml_buffer, xml_BufferPosition + 1, xml_AppendLength) = CStr(xml_Append)
        xml_BufferPosition = xml_BufferPosition + xml_AppendLength
    End If
End Sub

Private Function xml_BufferToString(ByRef xml_buffer As String, ByVal xml_BufferPosition As Long) As String
    If xml_BufferPosition > 0 Then
        xml_BufferToString = VBA.Left$(xml_buffer, xml_BufferPosition)
    End If
End Function

''
' VBA-UTC v1.0.6
' (c) Tim Hall - https://github.com/VBA-tools/VBA-UtcConverter
'
' UTC/ISO 8601 Converter for VBA
'
' Errors:
' 10011 - UTC parsing error
' 10012 - UTC conversion error
' 10013 - ISO 8601 parsing error
' 10014 - ISO 8601 conversion error
' 10015 - Unix parsing error
' 10016 - Uunix conversion error
'
' @module UtcConverter
' @author tim.hall.engr@gmail.com
' @license MIT (http://www.opensource.org/licenses/mit-license.php)
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '

' (Declarations moved to top)

' ============================================= '
' Public Methods
' ============================================= '

''
' Parse UTC date to local date
'
' @method ParseUtc
' @param {Date} UtcDate
' @return {Date} Local date
' @throws 10011 - UTC parsing error
''
Public Function ParseUtc(utc_UtcDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ParseUtc = utc_ConvertDate(utc_UtcDate)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_LocalDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_SystemTimeToTzSpecificLocalTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_UtcDate), utc_LocalDate

    ParseUtc = utc_SystemTimeToDate(utc_LocalDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10011, "UtcConverter.ParseUtc", "UTC parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to UTC date
'
' @method ConvertToUrc
' @param {Date} utc_LocalDate
' @return {Date} UTC date
' @throws 10012 - UTC conversion error
''
Public Function ConvertToUtc(utc_LocalDate As Date) As Date
    On Error GoTo utc_ErrorHandling

#If Mac Then
    ConvertToUtc = utc_ConvertDate(utc_LocalDate, utc_ConvertToUtc:=True)
#Else
    Dim utc_TimeZoneInfo As utc_TIME_ZONE_INFORMATION
    Dim utc_UtcDate As utc_SYSTEMTIME

    utc_GetTimeZoneInformation utc_TimeZoneInfo
    utc_TzSpecificLocalTimeToSystemTime utc_TimeZoneInfo, utc_DateToSystemTime(utc_LocalDate), utc_UtcDate

    ConvertToUtc = utc_SystemTimeToDate(utc_UtcDate)
#End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10012, "UtcConverter.ConvertToUtc", "UTC conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse ISO 8601 date string to local date
'
' @method ParseIso
' @param {Date} utc_IsoString
' @return {Date} Local date
' @throws 10013 - ISO 8601 parsing error
''
Public Function ParseIso(utc_IsoString As String) As Date
    On Error GoTo utc_ErrorHandling

    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String
    Dim utc_OffsetIndex As Long
    Dim utc_HasOffset As Boolean
    Dim utc_NegativeOffset As Boolean
    Dim utc_OffsetParts() As String
    Dim utc_Offset As Date

    utc_Parts = VBA.Split(utc_IsoString, "T")
    utc_DateParts = VBA.Split(utc_Parts(0), "-")
    ParseIso = VBA.DateSerial(VBA.CInt(utc_DateParts(0)), VBA.CInt(utc_DateParts(1)), VBA.CInt(utc_DateParts(2)))

    If UBound(utc_Parts) > 0 Then
        If VBA.InStr(utc_Parts(1), "Z") Then
            utc_TimeParts = VBA.Split(VBA.Replace(utc_Parts(1), "Z", vbNullString), ":")
        Else
            utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "+")
            If utc_OffsetIndex = 0 Then
                utc_NegativeOffset = True
                utc_OffsetIndex = VBA.InStr(1, utc_Parts(1), "-")
            End If

            If utc_OffsetIndex > 0 Then
                utc_HasOffset = True
                utc_TimeParts = VBA.Split(VBA.Left$(utc_Parts(1), utc_OffsetIndex - 1), ":")
                utc_OffsetParts = VBA.Split(VBA.Right$(utc_Parts(1), VBA.Len(utc_Parts(1)) - utc_OffsetIndex), ":")

                Select Case UBound(utc_OffsetParts)
                Case 0
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), 0, 0)
                Case 1
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), 0)
                Case 2
                    ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
                    utc_Offset = TimeSerial(VBA.CInt(utc_OffsetParts(0)), VBA.CInt(utc_OffsetParts(1)), Int(VBA.Val(utc_OffsetParts(2))))
                End Select

                If utc_NegativeOffset Then: utc_Offset = -utc_Offset
            Else
                utc_TimeParts = VBA.Split(utc_Parts(1), ":")
            End If
        End If

        Select Case UBound(utc_TimeParts)
        Case 0
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), 0, 0)
        Case 1
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), 0)
        Case 2
            ' VBA.Val does not use regional settings, use for seconds to avoid decimal/comma issues
            ParseIso = ParseIso + VBA.TimeSerial(VBA.CInt(utc_TimeParts(0)), VBA.CInt(utc_TimeParts(1)), Int(VBA.Val(utc_TimeParts(2))))
        End Select

        ParseIso = ParseUtc(ParseIso)

        If utc_HasOffset Then
            ParseIso = ParseIso - utc_Offset
        End If
    End If

    Exit Function

utc_ErrorHandling:
    Err.Raise 10013, "UtcConverter.ParseIso", "ISO 8601 parsing error for " & utc_IsoString & ": " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to ISO 8601 string
'
' @method ConvertToIso
' @param {Date} utc_LocalDate
' @return {Date} ISO 8601 string
' @throws 10014 - ISO 8601 conversion error
''
Public Function ConvertToIso(utc_LocalDate As Date) As String
    On Error GoTo utc_ErrorHandling

    ConvertToIso = VBA.Format$(ConvertToUtc(utc_LocalDate), "yyyy-mm-ddTHH:mm:ss.000Z")

    Exit Function

utc_ErrorHandling:
    Err.Raise 10014, "UtcConverter.ConvertToIso", "ISO 8601 conversion error: " & Err.Number & " - " & Err.Description
End Function

''
' Parse Unix timestamp to local date.
'
' @method ParseUnix
' @param {Long} Unix timestamp
' @return {Date} Local date
' @throws 10015 - Unix parsing error
''
Public Function ParseUnix(UnixDate As Long) As Date
    On Error GoTo utc_ErrorHandling
    
    ParseUnix = ParseUtc(DateAdd("s", UnixDate, "1/1/1970 00:00:00"))
    
    Exit Function

utc_ErrorHandling:
    Err.Raise 10015, "UtcConverter.ParseUnix", "Unix parsing error: " & Err.Number & " - " & Err.Description
End Function

''
' Convert local date to unix timestamp.
'
' @method ConvertToUnix
' @param {Date} LocalDate
' @return {String} Unix timestamp
' @throws 10016 - Unix conversion error
''
Public Function ConvertToUnix(LocalDate As Date) As Long
    On Error GoTo utc_ErrorHandling
    
    ConvertToUnix = VBA.DateDiff("s", "1/1/1970", ConvertToUtc(LocalDate))
    
    Exit Function

utc_ErrorHandling:
    Err.Raise 10016, "UtcConverter.ConvertToUnix", "Unix conversion error: " & Err.Number & " - " & Err.Description
End Function

' ============================================= '
' Private Functions
' ============================================= '

#If Mac Then

Private Function utc_ConvertDate(utc_Value As Date, Optional utc_ConvertToUtc As Boolean = False) As Date
    Dim utc_ShellCommand As String
    Dim utc_Result As utc_ShellResult
    Dim utc_Parts() As String
    Dim utc_DateParts() As String
    Dim utc_TimeParts() As String

    If utc_ConvertToUtc Then
        utc_ShellCommand = "date -ur `date -jf '%Y-%m-%d %H:%M:%S' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & "' " & _
            " +'%s'` +'%Y-%m-%d %H:%M:%S'"
    Else
        utc_ShellCommand = "date -jf '%Y-%m-%d %H:%M:%S %z' " & _
            "'" & VBA.Format$(utc_Value, "yyyy-mm-dd HH:mm:ss") & " +0000' " & _
            "+'%Y-%m-%d %H:%M:%S'"
    End If

    utc_Result = utc_ExecuteInShell(utc_ShellCommand)

    If utc_Result.utc_Output = vbNullString Then
        Err.Raise 10015, "UtcConverter.utc_ConvertDate", "'date' command failed"
    Else
        utc_Parts = VBA.Split(utc_Result.utc_Output, " ")
        utc_DateParts = VBA.Split(utc_Parts(0), "-")
        utc_TimeParts = VBA.Split(utc_Parts(1), ":")

        utc_ConvertDate = DateSerial(utc_DateParts(0), utc_DateParts(1), utc_DateParts(2)) + _
            TimeSerial(utc_TimeParts(0), utc_TimeParts(1), utc_TimeParts(2))
    End If
End Function

Private Function utc_ExecuteInShell(utc_ShellCommand As String) As utc_ShellResult
#If VBA7 Then
    Dim utc_File As LongPtr
    Dim utc_Read As LongPtr
#Else
    Dim utc_File As Long
    Dim utc_Read As Long
#End If

    Dim utc_Chunk As String

    On Error GoTo utc_ErrorHandling
    utc_File = utc_popen(utc_ShellCommand, "r")

    If utc_File = 0 Then: Exit Function

    Do While utc_feof(utc_File) = 0
        utc_Chunk = VBA.Space$(50)
        utc_Read = CLng(utc_fread(utc_Chunk, 1, Len(utc_Chunk) - 1, utc_File))
        If utc_Read > 0 Then
            utc_Chunk = VBA.Left$(utc_Chunk, CLng(utc_Read))
            utc_ExecuteInShell.utc_Output = utc_ExecuteInShell.utc_Output & utc_Chunk
        End If
    Loop

utc_ErrorHandling:
    utc_ExecuteInShell.utc_ExitCode = CLng(utc_pclose(utc_File))
End Function

#Else

Private Function utc_DateToSystemTime(utc_Value As Date) As utc_SYSTEMTIME
    utc_DateToSystemTime.utc_wYear = VBA.Year(utc_Value)
    utc_DateToSystemTime.utc_wMonth = VBA.Month(utc_Value)
    utc_DateToSystemTime.utc_wDay = VBA.Day(utc_Value)
    utc_DateToSystemTime.utc_wHour = VBA.Hour(utc_Value)
    utc_DateToSystemTime.utc_wMinute = VBA.Minute(utc_Value)
    utc_DateToSystemTime.utc_wSecond = VBA.Second(utc_Value)
    utc_DateToSystemTime.utc_wMilliseconds = 0
End Function

Private Function utc_SystemTimeToDate(utc_Value As utc_SYSTEMTIME) As Date
    utc_SystemTimeToDate = DateSerial(utc_Value.utc_wYear, utc_Value.utc_wMonth, utc_Value.utc_wDay) + _
        TimeSerial(utc_Value.utc_wHour, utc_Value.utc_wMinute, utc_Value.utc_wSecond)
End Function

#End If

