# Doctrack
Tool to insert tracking pixels into Office Open XML documents.
## Features
* Insert tracking pixels into Office Open XML documents (Word and Excel)
* Inject template URL (aka Remote Template Injection)
* Inspect external target URLs and metadata
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
Tool to insert tracking pixels into Office Open XML documents.
Copyright (C) 2021 doctrack

  -i, --input       Input filename.
  -o, --output      Output filename.
  -m, --metadata    Metadata to supply (json file).
  -u, --url         URL to insert.
  -e, --template    (Default: false) If set, enables template URL injection.
  -s, --inspect     (Default: false) Inspect external targets.
  --help            Display this help screen.
```
Insert tracking pixel and change document metadata:
```cmd
$ doctrack -i test.docx -o test.docx --metadata metadata.json --url http://test.url/image.png
```
Insert remote template URL (aka Remote Template Injection):
```cmd
$ doctrack -i test.docx -o test.docx --url http://test.url/template.dotm --template
```
Inspect external target URLs and metadata:
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
```