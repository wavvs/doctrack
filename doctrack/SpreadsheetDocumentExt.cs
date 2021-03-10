using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Linq;

namespace doctrack
{
    public static class SpreadsheetDocumentExt
    {
        // TODO: cleanup?
        public static void InsertTrackingURI(this SpreadsheetDocument workbook, string url)
        {
            WorkbookPart wbPart = workbook.WorkbookPart;
            var uri = new System.Uri(url);
            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().First();
            WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            DrawingsPart drawingsPart;
            if (wsPart.DrawingsPart == null)
            {
                drawingsPart = wsPart.AddNewPart<DrawingsPart>();
            }
            else
            {
                drawingsPart = wsPart.DrawingsPart;
            }
            var extRel = drawingsPart.AddExternalRelationship("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", uri);
            uint id;
            if (drawingsPart.WorksheetDrawing == null)
            {
                id = 1;
            }    
            else
            {
                id = Convert.ToUInt32(drawingsPart.WorksheetDrawing.Count()) + 1;
            }
            Random r = new Random();
            int rInt = r.Next(300, 2000);
            if (wsPart.Worksheet.Elements<Drawing>().Count() == 0)
            {
                GenerateDrawingsPart1Content(drawingsPart, extRel.Id, id, rInt, 0, rInt, 0, false, 0, 0);
            }
            else
            {
                GenerateDrawingsPart1Content(drawingsPart, extRel.Id, id, rInt, 0, rInt, 0, true, 0, 0);
            }
 
            wsPart.GetIdOfPart(drawingsPart);
            List<string> drawingsList = new List<string>();
            foreach (var drawing in wsPart.Worksheet.Elements<Drawing>())
            {
                drawingsList.Add(drawing.Id);
            }
            string drPartId = wsPart.GetIdOfPart(drawingsPart);
            Drawing drawingNew = new Drawing() { Id = drPartId };
            if (!drawingsList.Contains(drPartId))
            {
                wsPart.Worksheet.Append(drawingNew);
            }
                

        }
        private static void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1, string relId, uint id, int startRowIndex, int startColumnIndex, int endRowIndex, int endColumnIndex, bool appendToDrawing, int width, int height)
        {
            xdr.WorksheetDrawing worksheetDrawing1;
            if (!appendToDrawing)
            {
                worksheetDrawing1 = new xdr.WorksheetDrawing();
                worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            }    
            else
            {
                worksheetDrawing1 = drawingsPart1.WorksheetDrawing;
            }
                
            xdr.TwoCellAnchor twoCellAnchor1 = new xdr.TwoCellAnchor() { EditAs = xdr.EditAsValues.OneCell };
            xdr.FromMarker fromMarker1 = new xdr.FromMarker();
            xdr.ColumnId columnId1 = new xdr.ColumnId();
            columnId1.Text = startColumnIndex.ToString();
            xdr.ColumnOffset columnOffset1 = new xdr.ColumnOffset();
            columnOffset1.Text = "0";
            xdr.RowId rowId1 = new xdr.RowId();
            rowId1.Text = startRowIndex.ToString();
            xdr.RowOffset rowOffset1 = new xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            xdr.ToMarker toMarker1 = new xdr.ToMarker();
            xdr.ColumnId columnId2 = new xdr.ColumnId();
            columnId2.Text = endColumnIndex.ToString();
            xdr.ColumnOffset columnOffset2 = new xdr.ColumnOffset();
            columnOffset2.Text = "0";
            xdr.RowId rowId2 = new xdr.RowId();
            rowId2.Text = endRowIndex.ToString();
            xdr.RowOffset rowOffset2 = new xdr.RowOffset();
            rowOffset2.Text = "0";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            xdr.Picture picture1 = new xdr.Picture();
            xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new xdr.NonVisualPictureProperties();
            xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new xdr.NonVisualDrawingProperties()
            {
                Id = (UInt32Value)id,
                Name = Guid.NewGuid().ToString()
            };

            xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);
            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            xdr.BlipFill blipFill1 = new xdr.BlipFill();

            A.Blip blip1 = new A.Blip()
            {
                Link = relId,
            };

            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            xdr.ShapeProperties shapeProperties1 = new xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents()
            {
                Cx = (int)Math.Round((decimal)width * 9525),
                Cy = (int)Math.Round((decimal)height * 9525)

            };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            xdr.ClientData clientData1 = new xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);
            worksheetDrawing1.Append(twoCellAnchor1);
            if (!appendToDrawing)
            {
                drawingsPart1.WorksheetDrawing = worksheetDrawing1;
            }
        }
    }

}
