# VBA-XMLConverter

XML conversion and parsing for VBA (Excel, Access, and other Office applications).

Tested in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+. 

- For Windows-only support, include a reference to "Microsoft Scripting Runtime"
- For Mac support or to skip adding a reference, include [VBA-Dictionary](https://github.com/timhall/VBA-Dictionary).

# Example

```VB.net
Dim XML As Object
Set XML = XMLConverter.ParseXML( _
  "<?xml version="1.0"?>" & _
  "<messages>" & _
    "<message id="1" date="2014-1-1">" & _
      "<from><name>Tim Hall</name></from>" & _
      "<body>Howdy!</body>" & _
    "</message>" & _
  "</messages>" _
)

' XML("nodeName") -> messages
' XML("childNodes")(1)("attributes")("id") -> 1
' XML("childNodes")(1)("childNodes")(2)("text") -> "Howdy!"

Dim SearchResults As Collection
Set SearchResults = XMLConverter.QueryByXPath(XML, "/messages/message[1]/body")
' SearchResults(1)("text") -> "Howdy!"

Debug.Print XMLConverter.ConvertToXML(XML)
' -> "<?xml version="1.0"?><messages>...</messages>"
```
