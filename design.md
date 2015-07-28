# Goals

- Generally match `MSXML2.DOMDocument`
- Utilize `Dictionary` and `Collection` to add no new classes
- Xml = ConvertToXml(ParseXml(Xml))

# Parsing Components

1. prolog: `<? ... ?>`
2. doctype: `<!doctype ... [...]>`
3. documentElement: Root element
4. Element: 

```
#document (Element)
  prolog: (Element)
  doctype: (Element)
  documentElement: (Element)
  nodeName: "#document"
  childNodes:
  - (prolog)
  - (doctype)
  - (documentElement)
  attributes: (empty)
  text: ""
  xml: "..."
```

# Element

- nodeName `String`
- childNodes `Collection`
- attributes `Dictionary`
- text `String`
- xml `String`

```html
<messages name="Tim"><message id="1">A</message><message id="2">B</message></messages>
^         ^    ^    ^ ^              ^ ^         ^                ^         ^
a         b    c    d e              f g         h                i         j
```

`parseElement`

- a: `<`, nodeName = "messages"
- b: `non->`, attribute, key = "name"
- c: `=`, value = "Tim"
- d: `>`, close opening tag, check for childNodes/text/void next
- e: `< + non-/`, opening tag, childNodes + parseElement
- f: `non-<`, text = "A"
- g: `</`, closing tag, find > and exit parseElement
- h: `< + non-/`, another opening tag, childNodes + parseElement
- i: `</`, closing tag, find > and exit parseElement
- j: `</`, closing tag, find > and exit original parseElement

```
Look for space -> attributes
Look for > -> end of opening

If void element, look for immediate </...> or <
Otherwise:
  Look for immediate < -> childNodes -> parseElement
  Look for ... -> text
  Look for immediate </ -> close element
```

TODO Handle comments

`parseAttribute -> Array(Key, Value)`

`createElement(nodeName, childNodes, attributes, text, xml) -> Dictionary`

Helper for loading values into `Dictionary`

# Process

1. Parse prolog into Element
2. Parse doctype into Element
3. Use `parseElement` starting after doctype for documentElement
4. Create #document element and add prolog, doctype, and documentElement
