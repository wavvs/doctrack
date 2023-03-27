using System;
using System.IO;
using CommandLine;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml.Packaging;


namespace doctrack
{
    class Program
    {   
        class Options
        {
            [Option('i', "input", HelpText = "Input filename. If doesn't exist, new file is created.")]
            public string Input { get; set; }

            [Option('o', "output", HelpText = "Output filename. If not set, document is saved as --input file.")]
            public string Output { get; set; }

            [Option('m', "metadata", HelpText = "Metadata to supply (JSON file).")]
            public string Metadata { get; set; }
            
            [Option('u', "url", HelpText = "URL to insert.")]
            public string Url { get; set; }

            [Option('e', "template", Default = false, HelpText = "If set, enables template URL injection.")]
            public bool Template { get; set; }

            [Option('s', "inspect", Default = false, HelpText = "Inspect document.")]
            public bool Inspect { get; set; }

            [Option('c', "custom-part", HelpText = "Insert a Custom XML part (XML file)")]
            public string CustomPart { get; set; }
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
                        h.Heading = "Tool to manipulate and weaponize Office Open XML documents.";
                        h.Copyright = "";
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
                var isFileExist = File.Exists(opts.Input);

                var documentType = Path.GetExtension(opts.Input);
                switch (documentType)
                {
                    case ".docx":
                    case ".docm":
                    case ".dotm":
                    case ".dotx":
                        if (isFileExist)
                        {
                            using (var document = WordprocessingDocument.Open(opts.Input, false))
                            {
                                package = document.Clone();
                            }
                        } 
                        else
                        {
                            package = WordprocessingDocumentExt.Create(opts.Input, documentType);
                            if (package is null)
                            {
                                throw new OpenXmlPackageException();
                            }
                        }
                        break;
                    case ".xlsx":
                    case ".xlsm":
                    case ".xltm":
                    case ".xltx":
                        if (isFileExist)
                        {
                            using (var document = SpreadsheetDocument.Open(opts.Input, false))
                            {
                                package = document.Clone();
                            }
                        }
                        else
                        {
                            package = SpreadsheetDocumentExt.Create(opts.Input, documentType);
                            if (package is null)
                            {
                                throw new OpenXmlPackageException();
                            }
                        }
                        break;
                    default:
                        throw new OpenXmlPackageException();
                }

                if (opts.Inspect) return Utils.RunInspect(package);


                if (File.Exists(opts.Metadata))
                {
                    var obj = JObject.Parse(File.ReadAllText(opts.Metadata));
                    Utils.ModifyMetadata(package, obj);
                }

                if (!string.IsNullOrEmpty(opts.Url))
                {
                    if (opts.Template)
                    {
                        if (package is WordprocessingDocument w)
                        {
                           w.InsertTemplateURI(opts.Url);
                        }
                        else
                        {
                            Console.Error.WriteLine("[Error] Not supported.");
                            return 1;
                        }
                    }
                    else
                    {
                        if (package is WordprocessingDocument w)
                        {
                            w.InsertTrackingURI(opts.Url);
                        }
                        else if (package is SpreadsheetDocument s)
                        {
                            s.InsertTrackingURI(opts.Url);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(opts.CustomPart))
                {
                    if (!File.Exists(opts.CustomPart))
                    {
                        Console.Error.WriteLine("[Error] Specify XML file.");
                        return 1;
                    }

                    using (FileStream stream = new FileStream(opts.CustomPart, FileMode.Open))
                    {
                        if (package is WordprocessingDocument w)
                        {
                            w.AddCustomPart(stream);
                        } 
                        else if (package is SpreadsheetDocument s)
                        {
                            s.AddCustomPart(stream);
                        }
                    }
                }

                if (string.IsNullOrEmpty(opts.Output))
                {
                    opts.Output = opts.Input;
                }

                package.SaveAs(opts.Output);
                package.Close();
               
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
    }
}
