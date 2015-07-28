# VBA-XMLConverter

__Status__: _Incomplete, Under Development_

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

Debug.Print XML("documentElement")("nodeName") ' -> "messages"
Debug.Print XML("documentElement")("childNodes")(1)("attributes")("id") ' -> "1"
Debug.Print XML("documentElement")("childNodes")(1)("childNodes")(2)("text") ' -> "Howdy!"

Debug.Print XMLConverter.ConvertToXML(XML)
' -> "<?xml version="1.0"?><messages>...</messages>"
```
