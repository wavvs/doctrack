using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Schema;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using System.IO.Packaging;

namespace doctrack
{
    static class Utils
    {
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
    }
}
