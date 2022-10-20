using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;


namespace doctrack
{
    public static class WordprocessingDocumentExt
    {

        public static void InsertTemplateURI(this WordprocessingDocument document, string url)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            var uri = new Uri(url);
            DocumentSettingsPart documentSettingsPart = mainPart.DocumentSettingsPart;
            ExternalRelationship relationship = documentSettingsPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate", uri);
            documentSettingsPart.Settings.Append(
                new DocumentFormat.OpenXml.Wordprocessing.AttachedTemplate() { Id = relationship.Id });
        }
        public static void InsertTrackingURI(this WordprocessingDocument document, string url)
        {
            MainDocumentPart mainPart = document.MainDocumentPart;
            var uri = new System.Uri(url);
            var extRel = mainPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            var docProp = document.MainDocumentPart.Document.Descendants<DW.DocProperties>();
            uint id;
            if (docProp.Count() == 0)
                id = 0;
            else
                id = docProp.Max(element => element.Id.Value);
            var element = GetPictureElement(extRel.Id, Guid.NewGuid().ToString(), id + 1, 0, 0);
            document.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }

        public static void AddCustomPart(this WordprocessingDocument document, Stream xml)
        {
            MainDocumentPart mainDocumentPart = document.MainDocumentPart;
            CustomXmlPart customXmlPart = mainDocumentPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            CustomXmlPropertiesPart customXmlPropertiesPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
            customXmlPropertiesPart.DataStoreItem = Utils.GenerateCustomXMLProperties();
            customXmlPart.FeedData(xml);
        }

        private static Drawing GetPictureElement(string rId, string picname, uint id, int width, int height)
        {
            int emuCx = (int)Math.Round((decimal)width * 9525);
            int emuCy = (int)Math.Round((decimal)height * 9525);
            var element =
            new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = emuCx, Cy = emuCy },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = (UInt32Value)id,
                        Name = picname
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new A.GraphicFrameLocks() { NoChangeAspect = true }),
                    new A.Graphic(
                        new A.GraphicData(
                            new PIC.Picture(
                                new PIC.NonVisualPictureProperties(
                                    new PIC.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)id,
                                        Name = picname
                                    },
                                    new PIC.NonVisualPictureDrawingProperties()),
                                new PIC.BlipFill(
                                    new A.Blip(
                                        new A.BlipExtensionList(
                                            new A.BlipExtension()
                                            {
                                                Uri =
                                                "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            })
                                    )
                                    {
                                        Link = rId,
                                        CompressionState = A.BlipCompressionValues.Print
                                    },
                                    new A.Stretch(
                                        new A.FillRectangle())),
                                new PIC.ShapeProperties(
                                    new A.Transform2D(
                                        new A.Offset() { X = 0L, Y = 0L },
                                        new A.Extents() { Cx = emuCx, Cy = emuCy }),
                                    new A.PresetGeometry(
                                        new A.AdjustValueList()
                                    )
                                    { Preset = A.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                });

            return element;
        }
    }
}
