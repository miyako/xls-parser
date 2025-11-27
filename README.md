### Dependencies and Licensing

* the source code of this CLI tool is licensed under the MIT license.
* see [libxls](https://github.com/libxls/libxls/blob/dev/LICENSE) for the licensing of **libxls** (BSD).
 
# xls-parser
CLI tool to extract text from XLS

```
text extractor for xls documents

 -i path    : document to parse
 -o path    : text output (default=stdout)
 -          : use stdin for input
 -r         : raw text output (default=json)
 -c charset : charset (default=iso-8859-1)
```

## JSON

|Property|Level|Type|Description|
|-|-|-|-|
|document|0|||
|document.type|0|Text||
|document.pages|0|Array|=sheets|
|document.pages[].meta|1| Object ||
|document.pages[].meta.name|1| Text |sheet name|
|document.pages[].paragraphs|1|Array|=rows|
|document.pages[].paragraphs[].values|2|Array|=cells|
|document.pages[].paragraphs[].text|2|Text|JSON representation of .values|
