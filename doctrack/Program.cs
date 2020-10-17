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

            [Option('m', "metadata", HelpText = "Metadata to supply (json file)")]
            public string Metadata { get; set; }
            
            [Option('u', "url", HelpText = "URL to insert.")]
            public string Url { get; set; }

            [Option('e', "template", Default = false, HelpText = "If set, enables template URL injection.")]
            public bool Template { get; set; }

            [Option('t', "type", HelpText = "Document type. If --input is not specified, creates new document and saves as --output.")]
            public string Type { get; set; }

            [Option('l', "list-types", Default = false, HelpText = "Lists available types for document creation.")]
            public bool ListTypes { get; set; }

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
                        return CommandLine.Text.HelpText.DefaultParsingErrorsHandler(result, h);
                    }, e => e, true);
                    Console.WriteLine(helpText);
                    return 1;
                });

            return 0;
        }   

        static int RunOptions(Options opts)
        {
            if (opts.ListTypes) return RunListTypes();

            try
            {
                OpenXmlPackage package;
                if (string.IsNullOrEmpty(opts.Input))
                {
                    Console.Error.WriteLine("#TODO");
                    return 1;
                }
                if (string.IsNullOrEmpty(opts.Type))
                {
                    Console.Error.WriteLine("[Error] Specify -t, --type.");
                    return 1;
                }
                switch (opts.Type)
                {
                    case "Document":
                    case "MacroEnabledDocument":
                    case "MacroEnabledTemplate":
                    case "Template":
                        package = WordprocessingDocument.Open(opts.Input, true, new OpenSettings() { AutoSave = false });
                        break;
                    case "Workbook":
                    case "MacroEnabledWorkbook":
                    case "MacroEnabledTemplateX":
                    case "TemplateX":
                        package = SpreadsheetDocument.Open(opts.Input, true, new OpenSettings() { AutoSave = false });
                        break;
                    default:
                        Console.Error.WriteLine("[Error] Specify correct document type, use --list-types to view types.");
                        return 1;
                }

                if (opts.Inspect) return RunInspect(package);
                if (File.Exists(opts.Metadata))
                {
                    var obj = JObject.Parse(File.ReadAllText(opts.Metadata));
                    Utils.ModifyMetadata(package, obj);
                }
                if (string.IsNullOrEmpty(opts.Url))
                {
                    Console.Error.WriteLine("[Error] Specify -u, --url.");
                    return 1;
                }
                if (opts.Template)
                {
                    WordprocessingDocument document = (WordprocessingDocument)package;
                    document.InsertTemplateURI(opts.Url);
                }
                else
                {
                    string name = package.GetType().Name;
                    if (name == "WordprocessingDocument")
                    {
                        WordprocessingDocument document = (WordprocessingDocument)package;
                        document.InsertTrackingURI(opts.Url);
                    }
                    else if (name == "SpreadsheetDocument")
                    {
                        Console.Error.WriteLine("#TODO");
                        return 1;
                    }
                }
                if (string.IsNullOrEmpty(opts.Output))
                {
                    Console.Error.WriteLine("[Error] Specify -o, --output.");
                    return 1;
                }

                if (opts.Input == opts.Output)
                {
                    if (OpenXmlPackage.CanSave)
                    {
                        package.Save();
                    }
                    else
                    {
                        var clone = package.Clone();
                        package.Close();
                        clone.SaveAs(opts.Output);
                        clone.Close();
                    }
                } 
                else
                {
                    package.SaveAs(opts.Output);
                    package.Close();
                }
            }
            catch (OpenXmlPackageException)
            {
                Console.Error.WriteLine("[Error] Document type mismatch.");
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
            Console.WriteLine("External targets:");
            foreach (var rel in rels)
            {
                Console.WriteLine(String.Format("Part: {0}\nID: {1}\nURI: {2}\n", rel.Container, rel.Id, rel.Uri));
            }
            return 0;
        }

        static int RunListTypes()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("Document              (*.docx)\n");
            sb.Append("MacroEnabledDocument  (*.docm)\n");
            sb.Append("MacroEnabledTemplate  (*.dotm)\n");
            sb.Append("Template              (*.dotx)\n");
            sb.Append("Workbook              (*.xlsx)\n");
            sb.Append("MacroEnabledWorkbook  (*.xlsm)\n");
            sb.Append("MacroEnabledTemplateX (*.xltm)\n");
            sb.Append("TemplateX             (*.xltx)\n");
            Console.Write(sb.ToString());
            return 0;
        }
    }
}
