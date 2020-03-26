using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using System;
using System.IO;
using System.Xml.Linq;

namespace office
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 3)
            {
                PrintUsage();
                Environment.Exit(0);
            }

            FileInfo templateDoc = new FileInfo(args[0]);
            if (!templateDoc.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[0]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo dataFile = new FileInfo(args[1]);
            if (!dataFile.Exists)
            {
                Console.WriteLine("Error, {0} does not exist.", args[1]);
                PrintUsage();
                Environment.Exit(0);
            }
            FileInfo assembledDoc = new FileInfo(args[2]);
            if (assembledDoc.Exists)
            {
                Console.WriteLine("Error, {0} exists.", args[2]);
                PrintUsage();
                Environment.Exit(0);
            }

            WmlDocument wmlDoc = new WmlDocument(templateDoc.FullName);
            XElement data = XElement.Load(dataFile.FullName);
            bool templateError;
            WmlDocument wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, data, out templateError);
            if (templateError)
            {
                Console.WriteLine("Errors in template.");
                Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
            }

            wmlAssembledDoc.SaveAs(assembledDoc.FullName);
        }

        static void PrintUsage()
        {
            Console.WriteLine("Usage: DocumentAssembler TemplateDocument.docx Data.xml AssembledDoc.docx");
        }

        static void Main1(string[] args)
        {
            GenerateDocument();
            Console.WriteLine("Hello World!");
        }

        public static void GenerateDocument()
        {
            string rootPath = @"C:\OfficeDocs";
            string xmlDataFile = @"test.xml";
            string templateDocument = @"temp.docx";
            string outputDocument = rootPath + @"\MyGeneratedDocument.docx";

            File.Copy(templateDocument, outputDocument);
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputDocument, true))
            {
                
                //get the main part of the document which contains CustomXMLParts
                MainDocumentPart mainPart = wordDoc.MainDocumentPart;

                //delete all CustomXMLParts in the document. If needed only specific CustomXMLParts can be deleted using the CustomXmlParts IEnumerable
                mainPart.DeleteParts<CustomXmlPart>(mainPart.CustomXmlParts);

                //add new CustomXMLPart with data from new XML file
                CustomXmlPart myXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (FileStream stream = new FileStream(xmlDataFile, FileMode.Open))
                {
                    myXmlPart.FeedData(stream);
                }
            }

        }
    }
}
