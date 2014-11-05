# VBA-XMLConverter

XML conversion and parsing for VBA (Excel, Access, and other Office applications).

Tested in Windows Excel 2013 and Excel for Mac 2011, but should apply to 2007+. 

- For Windows-only support, include a reference to "Microsoft Scripting Runtime"
- For Mac support or to skip adding a reference, include [VBA-Dictionary](https://github.com/timhall/VBA-Dictionary).

# Example

```VB
Dim XML As Object
Set XML = XMLConverter.ParseXML("...")

' ...

Debug.Print XMLConverter.ConvertToXML(XML)
' -> ...
```
