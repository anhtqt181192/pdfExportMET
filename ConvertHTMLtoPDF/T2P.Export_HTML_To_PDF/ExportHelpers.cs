

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.parser;
using iTextSharp.tool.xml.pipeline.css;
using iTextSharp.tool.xml.pipeline.end;
using iTextSharp.tool.xml.pipeline.html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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

            PdfPTable table = new PdfPTable(3);
            Font footerFont = new Font(Font.FontFamily.TIMES_ROMAN, 10);
            footerFont.Color = new BaseColor(51, 51, 51);
            Chunk para = new Chunk("   itelya GmbH & Co. KG", footerFont);

            PdfPCell Topcell = new PdfPCell();
            Topcell.AddElement(para);
            Topcell.BorderWidthBottom = 0f;
            Topcell.BorderWidthLeft = 2f;
            Topcell.BorderWidthRight = 0f;
            Topcell.BorderWidthTop = 0f;
            Topcell.BorderColor = new BaseColor(112, 173, 71);
            Topcell.FixedHeight = 20f;
            Topcell.HorizontalAlignment = Element.ALIGN_RIGHT;
            Topcell.VerticalAlignment = Element.ALIGN_TOP;

            table.AddCell(Topcell);
            table.AddCell(Topcell);
            table.AddCell(Topcell);

            table.AddCell(Topcell);
            table.AddCell(Topcell);
            table.AddCell(Topcell);

            table.AddCell(Topcell);
            table.AddCell(Topcell);
            table.AddCell(Topcell);

            using (var ms = new MemoryStream())
            {
                using (var stamper = new PdfStamper(pdfReader, ms))
                {
                    int PageCount = pdfReader.NumberOfPages;
                    for (int i = 1; i <= PageCount; i++)
                    {
                        PdfContentByte content = stamper.GetOverContent(i);
                        content.SetColorFill(new BaseColor(51, 51, 51));
                        Phrase phare = new Phrase(i.ToString());
                        
                        ColumnText ct = new ColumnText(content);
                        var rectangle = new Rectangle(0, 0, 550, 75);
                        ct.SetSimpleColumn(rectangle);
                        ct.AddElement(table);
                        ct.Go();

                        content.Rectangle(20, 20, 550, 75);
                        ColumnText.ShowTextAligned(content, Element.ALIGN_RIGHT, phare, 575, 20, 0);
                    }
                }
                bytes = ms.ToArray();
            }
            bytes.ToFile(docUrl);
            pdfReader.Close();
        }

        public static byte[] AddFooterToPDF(PdfWriter writer, Document doc, string html, string css)
        {
            byte[] bytes;
            var ms = new MemoryStream();
            var msCss = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(css));
            var msHtml = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(html));
            try
            {
                iTextSharp.tool.xml.XMLWorkerHelper.GetInstance().ParseXHtml(writer, doc, msHtml, msCss);
            }
            catch (Exception e)
            {
            }
            msCss.Close();
            msHtml.Close();
            bytes = ms.ToArray();
            ms.Close();
            return bytes;
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
