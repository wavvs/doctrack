using System;
using System.IO;
using CommandLine;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json.Schema;
using System.Text;
using System.IO.Packaging;

namespace doctrack
{
    class Program
    {   
        class Options
        {
            [Option('i', "input", HelpText = "Input filename.")]
            public string Input { get; set; }

            [Option('o', "output", HelpText = "Output filename.")]
            public string Output { get; set; }

            [Option('m', "metadata", HelpText = "Metadata to supply (json file).")]
            public string Metadata { get; set; }
            
            [Option('u', "url", HelpText = "URL to insert.")]
            public string Url { get; set; }

            [Option('e', "template", Default = false, HelpText = "If set, enables template URL injection.")]
            public bool Template { get; set; }

            [Option('s', "inspect", Default = false, HelpText = "Inspect external targets.")]
            public bool Inspect { get; set; }
        }
        static int Main(string[] args)
        {
            var parser = new Parser(with =>
            {
                with.EnableDashDash = true;
                with.HelpWriter = null;
                with.AutoVersion = false;
            });
            var result = parser.ParseArguments<Options>(args);
            return result.MapResult(
                (Options opts) => RunOptions(opts),
                errs => {
                    var helpText = CommandLine.Text.HelpText.AutoBuild(result, h =>
                    {
                        h.AdditionalNewLineAfterOption = false;
                        h.AddDashesToOption = true;
                        h.AutoVersion = false;
                        h.Heading = "Tool to insert tracking pixels into Office Open XML documents.";
                        return CommandLine.Text.HelpText.DefaultParsingErrorsHandler(result, h);
                    }, e => e, true);
                    Console.WriteLine(helpText);
                    return 1;
                });
        }   

        static int RunOptions(Options opts)
        {
            try
            {
                OpenXmlPackage package;
                if (!File.Exists(opts.Input))
                {
                    Console.Error.WriteLine("[Error] Specify -i,--input.");
                    return 1;
                }

                var documentType = Path.GetExtension(opts.Input);
                switch (documentType)
                {
                    case ".docx":
                    case ".docm":
                    case ".dotm":
                    case ".dotx":
                        using (var document = WordprocessingDocument.Open(opts.Input, false))
                        {
                            package = document.Clone();
                        }
                        break;
                    case ".xlsx":
                    case ".xlsm":
                    case ".xltm":
                    case ".xltx":
                        using (var document = SpreadsheetDocument.Open(opts.Input, false))
                        {
                            package = document.Clone();
                        }
                        break;
                    default:
                        throw new OpenXmlPackageException();
                }

                if (opts.Inspect) return RunInspect(package);

                if (File.Exists(opts.Metadata))
                {
                    var obj = JObject.Parse(File.ReadAllText(opts.Metadata));
                    Utils.ModifyMetadata(package, obj);
                }

                if (!string.IsNullOrEmpty(opts.Url))
                {
                    string name = package.GetType().Name;
                    // TODO: Is it possible to inject templates into xlsx?
                    if (opts.Template)
                    {
                        if (name == "WordprocessingDocument")
                        {
                            WordprocessingDocument document = (WordprocessingDocument)package;
                            document.InsertTemplateURI(opts.Url);
                        }
                        else
                        {
                            Console.Error.WriteLine("[Error] Not supported.");
                            return 1;
                        }
                    }
                    else
                    {
                        if (name == "WordprocessingDocument")
                        {
                            WordprocessingDocument document = (WordprocessingDocument)package;
                            document.InsertTrackingURI(opts.Url);
                        }
                        else if (name == "SpreadsheetDocument")
                        {
                            SpreadsheetDocument workbook = (SpreadsheetDocument)package;
                            workbook.InsertTrackingURI(opts.Url);
                        }
                    }
                }

                if (string.IsNullOrEmpty(opts.Output))
                {
                    Console.Error.WriteLine("[Error] Specify -o, --output.");
                    return 1;
                }

                package.SaveAs(opts.Output);
               
            }
            catch (OpenXmlPackageException)
            {
                Console.Error.WriteLine("[Error] Document type mismatch. Provide correct Office Open XML file.");
                return 1;
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("[Error] {0}", e.Message);
                return 1;
            }
            return 0;
        }

        static int RunInspect(OpenXmlPackage package)
        {
            var rels = Utils.InspectExternalRelationships(package);
            Console.WriteLine("[External targets]");
            foreach (var rel in rels)
            {
                Console.WriteLine(String.Format("Part: {0}, ID: {1}, URI: {2}", rel.SourceUri, rel.Id, rel.TargetUri));
            }

            Console.WriteLine("[Metadata]");
            var propInfo = package.PackageProperties.GetType().GetProperties();
            foreach (var info in propInfo)
            {
                Console.WriteLine("{0}: {1}", info.Name, info.GetValue(package.PackageProperties));
            }
            return 0;
        }
    }
}
