using System;
using System.Linq;
using System.Collections.Generic;
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
