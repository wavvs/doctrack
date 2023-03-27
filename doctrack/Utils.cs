using System;
using System.Linq;
using System.Xml;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Linq;


namespace doctrack
{
    static class Utils
    {
        // TODO: Cannot change metadata on files created by LibreOffice.
        public static void ModifyMetadata(OpenXmlPackage package, JObject metadata)
        {
            string[] props = {
                "Created",
                "LastPrinted",
                "Modified",
                "Category",
                "ContentStatus",
                "Creator",
                "ContentType",
                "Description",
                "Identifier",
                "Keywords",
                "LastModifiedBy",
                "Language",
                "Revision",
                "Subject",
                "Title",
                "Version"
            };

            try
            {
                foreach (var pair in metadata)
                {
                    var key = pair.Key.ToString();
                    var value = pair.Value.ToString();
                    if (!props.Contains(key))
                    {
                        continue;
                    }
                    var property = package.PackageProperties.GetType().GetProperty(key);
                    if (key == props[0] || key == props[1] || key == props[2])
                    {
                        var date = Convert.ToDateTime(value);
                        property.SetValue(package.PackageProperties, date);
                    }
                    else
                    {
                        property.SetValue(package.PackageProperties, value);
                    }
                }
            } 
            catch (FormatException)
            {
                throw;
            }
        }

        // TODO: /word/footnotes.xml shows useless info.
        public static List<PackageRelationship> InspectExternalRelationships(OpenXmlPackage package)
        {
            var packageRels = new List<PackageRelationship>();
            foreach (var part in package.Package.GetParts())
            {
                if (!part.Uri.ToString().EndsWith(".rels"))
                {
                    var rels = part.GetRelationships();
                    foreach (var rel in rels)
                    {
                        if (rel.TargetMode == TargetMode.External)
                        {
                            packageRels.Add(rel);
                        }
                    }
                }
            }
            return packageRels;
        }

        public static DataStoreItem GenerateCustomXMLProperties()
        {
            DataStoreItem dataStoreItem = new DataStoreItem() { ItemId = "{" + Guid.NewGuid().ToString() + "}" };
            dataStoreItem.AddNamespaceDeclaration("ds", "http://schemas.openxmlformats.org/officeDocument/2006/customXml");
            SchemaReferences schemaReferences = new SchemaReferences();
            dataStoreItem.Append(schemaReferences);
            return dataStoreItem;
        }

        public static void AddCoreFileProperties(CoreFilePropertiesPart part)
        {
            using (var writer = new XmlTextWriter(part.GetStream(FileMode.Create), Encoding.UTF8))
            {
                string currentTime = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ssK");
                writer.WriteRaw("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"+
                    "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "+ 
                    "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "+
                    "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">"+
                    "<dc:title></dc:title><dc:subject></dc:subject><dc:creator></dc:creator><cp:keywords></cp:keywords>"+
                    "<dc:description></dc:description><cp:lastModifiedBy></cp:lastModifiedBy><cp:revision>1</cp:revision>"+
                    $"<dcterms:created xsi:type=\"dcterms:W3CDTF\">{currentTime}</dcterms:created>"+
                    $"<dcterms:modified xsi:type=\"dcterms:W3CDTF\">{currentTime}</dcterms:modified></cp:coreProperties>");
                writer.Flush();
            }
        }

        public static int RunInspect(OpenXmlPackage package)
        {
            var rels = InspectExternalRelationships(package);
            Console.WriteLine("[External targets]");
            foreach (var rel in rels)
            {
                Console.WriteLine(String.Format("Part: {0}, ID: {1}, URI: {2}", rel.SourceUri, rel.Id, rel.TargetUri));
            }
            Console.WriteLine("\n[Metadata]");
            var propInfo = package.PackageProperties.GetType().GetProperties();
            foreach (var info in propInfo)
            {
                Console.WriteLine("{0}: {1}", info.Name, info.GetValue(package.PackageProperties));
            }

            Console.WriteLine("\n[CustomXML Parts]");
            foreach (var part in package.RootPart.Parts)
            {
                string path = part.OpenXmlPart.Uri.ToString().ToLower();
                if (path.Contains("customxml"))
                {
                    Console.WriteLine("Part: {0} ({1} bytes)", part.OpenXmlPart.Uri, 
                        part.OpenXmlPart.GetStream().Length);
                }
            }
            return 0;
        }
    }
}
