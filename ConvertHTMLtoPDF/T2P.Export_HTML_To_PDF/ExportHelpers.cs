

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.html;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using T2P.Export_HTML_To_PDF.Extentions;
using T2P.Export_HTML_To_PDF.Models;

namespace T2P.Export_HTML_To_PDF
{
    public static class ExportHelpers
    {
        public static void ToPDF(PDFTemplate template)
        {
            string html = System.IO.File.ReadAllText(template.url_html);
            string html_details = System.IO.File.ReadAllText(template.url_html_details);
            string footer = System.IO.File.ReadAllText(template.url_html_footer);
            string css = System.IO.File.ReadAllText(template.url_css);
            byte[] fontStream = null;
            
            if (!String.IsNullOrEmpty(template.url_font))
            {
                fontStream = System.IO.File.ReadAllBytes(template.url_font);
            }

            List<string> filePDF = new List<string>();
            filePDF.Add(template.dir + $"\\template_{DateTime.Now.Ticks}.pdf");
            filePDF.Add(template.dir + $"\\template_{DateTime.Now.Ticks}_details.pdf");
            string pdfResult = template.dir + "\\" + template.out_FileName;
            CovnertHTMLToPDF(html, css, footer).ToFile(filePDF[0]);
            CovnertHTMLToPDF(html_details, css, footer).ToFile(filePDF[1]);
            CombineMultiplePDFs(filePDF.ToArray(), pdfResult);
            RenderFooter(pdfResult);
            foreach(var item in filePDF)
            {
                File.Delete(item);
            }    
            return;
        }

        public static byte[] CovnertHTMLToPDF(string html, string css, string footer)
        {
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(PageSize.A4, 25, 25, 25, 75))
                {
                    using (var writer = PdfWriter.GetInstance(doc, ms))
                    {
                        doc.Open();
                        using (var msCss = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(css)))
                        {
                            using (var msHtml = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html)))
                            {
                                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msHtml, msCss);
                            }
                        }
                        doc.Close();
                    }
                }
                bytes = ms.ToArray();
            }
            return bytes;
        }

        public static void RenderFooter(string docUrl)
        {
            var bytes = ByteHelpers.GetByteFromFile(docUrl);
            PdfReader pdfReader = new PdfReader(bytes);

            Font footerFont = new Font(Font.FontFamily.TIMES_ROMAN, 9);
            footerFont.Color = new BaseColor(51, 51, 51);

            Dictionary<int, List<string>> rows = new Dictionary<int, List<string>>();

            List<string> listPara1 = new List<string>();
            listPara1.Add("itelya GmbH & Co. KG");
            listPara1.Add("Bahnhofstraße 8 • 97500 Ebelsbach");
            listPara1.Add("Amtsgericht Bamberg • HRA 11272");
            listPara1.Add("UST-ID Nr.: DE271950303");
            rows[0] = listPara1;

            List<string> listPara2 = new List<string>();
            listPara2.Add("Persönlich haftende Gesellschaft: itelya Verwaltungs-GmbH");
            listPara2.Add("Geschäftsführer: Ralf Maier");
            listPara2.Add("Sitz der Gesellschaft: Ebelsbach");
            listPara2.Add("Amtsgericht Bamberg • HRB 6728");
            rows[1] = listPara2;

            List<string> listPara3 = new List<string>();
            listPara3.Add("Bankv VR Bank Bamberg-Forchheim eG");
            listPara3.Add("Bankleitzahl: 763 9100 0 • Kto.- Nr. 1071 629 36");
            listPara3.Add("IBAN: DE76 7639 1000 0107 1629 36");
            listPara3.Add("BIC: GENODEF1FOH");
            rows[2] = listPara3;

            var table = GenerateTable(rows);

            using (var ms = new MemoryStream())
            {
                using (var stamper = new PdfStamper(pdfReader, ms))
                {
                    int PageCount = pdfReader.NumberOfPages;
                    for (int i = 1; i <= PageCount; i++)
                    {
                        PdfContentByte content = stamper.GetOverContent(i);
                        ColumnText ct = new ColumnText(content);
                        var rectangle = new Rectangle(0, 0, 600, 75);
                        ct.SetSimpleColumn(rectangle);
                        ct.AddElement(table);
                        ct.Go();
                        ColumnText.ShowTextAligned(content, Element.ALIGN_RIGHT, GenerateParagraph(new List<string> { i.ToString() }), 575, 20, 0);
                    }
                }
                bytes = ms.ToArray();
            }
            bytes.ToFile(docUrl);
            pdfReader.Close();
        }

        public static Paragraph GenerateParagraph(List<string> str)
        {
            Font footerFont = new Font(Font.FontFamily.TIMES_ROMAN, 9);
            footerFont.Color = new BaseColor(71, 71, 71);

            Paragraph elements = new Paragraph();
            string last = str.Last();
            foreach (var item in str)
            {
                elements.Add(new Paragraph(item, footerFont));
                if(!item.Equals(last))
                    elements.Add(new Paragraph(Environment.NewLine));
            }
            return elements;
        }

        public static PdfPCell GenerateCell(Paragraph elements)
        {
            PdfPCell cell = new PdfPCell(elements);
            cell.BorderColor = new BaseColor(255, 255, 255);
            cell.BorderWidthLeft = 2f;
            cell.BorderColorLeft = new BaseColor(112, 173, 71);
            cell.HorizontalAlignment = Element.ALIGN_LEFT;
            cell.VerticalAlignment = Element.ALIGN_TOP;
            cell.Padding = 5;
            cell.UseAscender = true;
            return cell;
        }

        public static PdfPTable GenerateTable(Dictionary<int, List<string>> rows)
        {
            PdfPTable table = new PdfPTable(3);
            table.AddCell(GenerateCell(GenerateParagraph(rows[0])));
            table.AddCell(GenerateCell(GenerateParagraph(rows[1])));
            table.AddCell(GenerateCell(GenerateParagraph(rows[2])));
            table.WidthPercentage = 95;
            int[] widths = new int[] { 24, 39, 32 };
            table.SetWidths(widths);
            return table;
        }

        public static void CombineMultiplePDFs(string[] fileNames, string outFile)
        {
            Document document = new Document();
            using (FileStream newFileStream = new FileStream(outFile, FileMode.Create))
            {
                PdfCopy writer = new PdfCopy(document, newFileStream);
                if (writer == null)
                {
                    return;
                }
                document.Open();
                foreach (string fileName in fileNames)
                {
                    PdfReader reader = new PdfReader(fileName);
                    reader.ConsolidateNamedDestinations();
                    for (int i = 1; i <= reader.NumberOfPages; i++)
                    {
                        PdfImportedPage page = writer.GetImportedPage(reader, i);
                        writer.AddPage(page);
                    }
                    PRAcroForm form = reader.AcroForm;
                    if (form != null)
                    {
                        writer.CopyDocumentFields(reader);
                    }
                    reader.Close();
                }
                writer.Close();
                document.Close();
            }
        }
    }
}
