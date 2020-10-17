using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Schema;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;

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

        public static List<ExternalRelationship> InspectExternalRelationships(OpenXmlPackage package)
        {
            List<ExternalRelationship> extRels = new List<ExternalRelationship>();
            foreach (var part in package.Parts)
            {
                foreach (var rels in part.OpenXmlPart.ExternalRelationships)
                {
                    extRels.Add(rels);
                }
            }

            foreach (var part in package.RootPart.Parts)
            {
                foreach (var rels in part.OpenXmlPart.ExternalRelationships)
                {
                    extRels.Add(rels);
                }
            }
            return extRels;
        }
    }
}
