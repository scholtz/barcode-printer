using iText.Barcodes;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Xobject;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.IO;

namespace BarcodePrinter
{
    class Settings
    {
        public int OffsetBranch { get; set; } = 1000000000;
        public int OffsetIter { get; set; } = 1;
        public int Columns { get; set; } = 3;
        public int Count { get; set; } = 100;
        public string SaveToFile { get; set; } = "results/barcodes.pdf";
        public float Height { get; set; } = 50f;
        public float Width { get; set; } = 200f;
        public float CellPaddingTop { get; set; } = 10f;
        public float CellPaddingRight { get; set; } = 10f;
        public float CellPaddingLeft { get; set; } = 30f;
        public float CellPaddingBottom { get; set; } = 10f;
        public float PageMarginTop { get; set; } = 45f;
        public float PageMarginRight { get; set; } = 45f;
        public float PageMarginLeft { get; set; } = 45f;
        public float PageMarginBottom { get; set; } = 45f;
    }
    class Program
    {

        public static void Main(String[] args)
        {
            var settings = JsonConvert.DeserializeObject<Settings>(File.ReadAllText("appsettings.json"));
            FileInfo file = new FileInfo(settings.SaveToFile);
            file.Directory.Create();
            new Program().ManipulatePdf(settings);
        }

        private void ManipulatePdf(Settings settings)
        {
            PdfDocument pdfDoc = new PdfDocument(new PdfWriter(settings.SaveToFile));
            Document doc = new Document(pdfDoc, iText.Kernel.Geom.PageSize.A4);

            doc.SetTopMargin(settings.PageMarginTop);
            doc.SetLeftMargin(settings.PageMarginLeft);
            doc.SetBottomMargin(settings.PageMarginBottom);
            doc.SetRightMargin(settings.PageMarginRight);
            Table table = new Table(UnitValue.CreatePercentArray(settings.Columns)).UseAllAvailableWidth();
            for (int i = settings.OffsetIter; i < settings.OffsetIter + settings.Count; i++)
            {
                table.AddCell(CreateBarcode(string.Format("{0:d8}", i + settings.OffsetBranch), pdfDoc, settings));
            }

            doc.Add(table);

            doc.SetMargins(settings.PageMarginTop, settings.PageMarginRight, settings.PageMarginBottom, settings.PageMarginLeft);
            doc.Close();
        }

        private static Cell CreateBarcode(string code, PdfDocument pdfDoc, Settings settings)
        {
            Barcode39 barcode = new Barcode39(pdfDoc);
            barcode.SetCodeType(Barcode39.ALIGN_CENTER);
            barcode.SetCode(code);
            barcode.SetBarHeight(settings.Height);
            barcode.FitWidth(settings.Width);

            // Create barcode object to put it to the cell as image
            PdfFormXObject barcodeObject = barcode.CreateFormXObject(null, null, pdfDoc);
            Cell cell = new Cell().Add(new Image(barcodeObject));
            cell.SetPaddingTop(settings.CellPaddingTop);
            cell.SetPaddingRight(settings.CellPaddingRight);
            cell.SetPaddingBottom(settings.CellPaddingBottom);
            cell.SetPaddingLeft(settings.CellPaddingLeft);

            cell.SetBorder(new iText.Layout.Borders.DottedBorder(iText.Kernel.Colors.ColorConstants.LIGHT_GRAY, 1));


            return cell;
        }
    }
}
