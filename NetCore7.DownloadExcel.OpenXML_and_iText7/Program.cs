using ClosedXML.Excel;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;
using Table = iText.Layout.Element.Table;
using iText.Kernel.Geom;
using Style = iText.Layout.Style;

namespace NetCore7.DownloadExcel.OpenXML_and_iText7
{
    internal class Program
    {
        readonly static string rootFolderPath = "D:\\recycle-bin\\";
        static void Main(string[] args)
        {
            Console.WriteLine("Running .Net Core 7 Console App!");

            Console.WriteLine("Directory created as "+ rootFolderPath);
            CreateDirectory(rootFolderPath);

            Console.WriteLine("*** Creating a file using CLosed XML ***");
            ClosedXML();
            Console.WriteLine("*** File Creation is completed using CLosed XML ***");

            Console.WriteLine("*** Creating a file using iText7 ***");
            iText7();
            Console.WriteLine("*** File Creation is completed using iText7 ***");

            Console.WriteLine("Program is completed");
        }

        protected static void ClosedXML()
        {
           string filename = GenerateFileName(".xlsx");
            string ExcelFile = rootFolderPath + filename;

            var wbook = new XLWorkbook();
            var ws = wbook.AddWorksheet("Sheet1");
            int xl_r = 1; // first row
            int xl_c = 3;
            ws.Cell(xl_r, xl_c).Value = "Closed XML Demo";
            ws.Cell(xl_r, xl_c).Style.Font.Bold = true;
            ws.Cell(xl_r, xl_c).Style.Font.FontSize = 32;
            ws.Cell(xl_r, xl_c).Style.Fill.BackgroundColor = XLColor.LightGreen;
            ws.Column(xl_c).AdjustToContents();

            xl_r = 3; // first row
            xl_c = 3;
            ws.Cell(xl_r, xl_c).Value = "Sample Text";
            ws.Cell(xl_r, xl_c).Style.Font.Bold = true;
            ws.Cell(xl_r, xl_c).Style.Fill.BackgroundColor = XLColor.LightGreen;

            ClosedXML.Excel.SaveOptions optSave = new ClosedXML.Excel.SaveOptions { EvaluateFormulasBeforeSaving = false, GenerateCalculationChain = false, ValidatePackage = false };
            wbook.SaveAs(ExcelFile, optSave);
        }
        protected static void iText7()
        {
            Document MasterDoc;
            PdfWriter MasterWriter;

            string filename = GenerateFileName(".pdf");
            //string outfile2 = rootFolderPath + filename;
            string outfile2 = "D:\\sqloutput\\demo.pdf";
            MasterWriter = new PdfWriter(outfile2);
            PdfDocument pdfDocument = new PdfDocument(MasterWriter);
            pdfDocument.SetDefaultPageSize(PageSize.A4);
            MasterDoc = new Document(pdfDocument);

            PdfFont font = PdfFontFactory.CreateFont(StandardFonts.COURIER_BOLD);
            PdfFont font2 = PdfFontFactory.CreateFont(StandardFonts.COURIER);
            PdfFont fWhite = PdfFontFactory.CreateFont(StandardFonts.COURIER);

            Style styleBold = new Style();
            styleBold.SetFont(font);
            styleBold.SetFontSize(9f);

            Style styleTitle = new Style();
            styleTitle.SetFont(font);
            styleTitle.SetFontSize(32f);
            styleTitle.SetFontColor(ColorConstants.RED);

            Style styleTitle2 = new Style();
            styleTitle2.SetFont(font);
            styleTitle2.SetFontSize(32f);
            styleTitle2.SetFontColor(ColorConstants.BLUE);

            Style styleNormal = new Style();
            styleNormal.SetFont(font2);
            styleNormal.SetFontSize(9f);

            Style styleWhite = new Style();
            styleWhite.SetFont(fWhite);
            styleWhite.SetFontSize(9f);
            styleWhite.SetFontColor(ColorConstants.WHITE);

            Table table1 = new Table(2, false);

            string ttText = "iText7";

            Paragraph ttPar = new Paragraph(ttText).AddStyle(styleTitle);

            Cell ttCell1 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).AddStyle(styleTitle).Add(ttPar)
                .SetBorder(Border.NO_BORDER)
                .SetMarginRight(0f)
                .SetWidth(200f)
                .SetHeight(50f)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                ;

            ttText = "DEMO";
            Paragraph ttPar2a = new Paragraph(ttText).AddStyle(styleTitle2);

            Cell ttCell2a = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).AddStyle(styleTitle2).Add(ttPar2a)
                .SetBorder(Border.NO_BORDER)
                .SetMarginRight(0f)
                .SetWidth(100f)
                .SetHeight(50f)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                ;

            table1.AddCell(ttCell1);
            table1.AddCell(ttCell2a);
            table1.SetHeight(60f);
            MasterDoc.Add(table1);

            // next paragraph
            Table table2 = new Table(1, false);

            ttText = "Sample Text";

            Paragraph ttPar2 = new Paragraph(ttText).AddStyle(styleNormal);

            Cell ttCell2 = new Cell(1, 1).SetTextAlignment(TextAlignment.LEFT).AddStyle(styleNormal).Add(ttPar2)
                .SetBorder(Border.NO_BORDER)
                .SetMarginRight(0f)
                .SetWidth(300f)
                .SetHeight(18f)
                .SetVerticalAlignment(VerticalAlignment.TOP)
                ;

            table2.AddCell(ttCell2);
            table2.SetHeight(20f);
            MasterDoc.Add(table2);

            MasterDoc.Close();
            MasterWriter.Close();
        }

        protected static string GenerateFileName(string ext)
        {
            string ddmmyyhhss = DateTime.Now.ToString("ddMMyyyyHHmmss");
            return ddmmyyhhss + ext;

        }

        protected static void CreateDirectory(string directory)
        {
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);

            }
        }
    }
}