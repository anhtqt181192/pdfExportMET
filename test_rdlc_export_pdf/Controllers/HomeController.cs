using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using Microsoft.Reporting.WebForms;
using Microsoft.ReportingServices.Diagnostics.Internal;
using Newtonsoft.Json;
using PdfSharp.Charting;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Web.Mvc;
using System.Web.UI;
using test_rdlc_export_pdf.App_Start;

namespace test_rdlc_export_pdf.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult HTMLToPDF()
        {
            try
            {
                string html = System.IO.File.ReadAllText(HttpContext.Server.MapPath("~\\Template\\template.html"));
                string css = System.IO.File.ReadAllText(HttpContext.Server.MapPath("~\\Template\\template.css"));
                byte[] fontStream = System.IO.File.ReadAllBytes(HttpContext.Server.MapPath("~\\Template\\Roboto\\Roboto-Regular.ttf"));
                var bytes = CovnertHTMLToPDF(html, css, fontStream);
                //ByteArrayToFile(bytes, HttpContext.Server.MapPath("~\\Template\\testfile.pdf"));
                Response.Buffer = true;
                Response.Clear();
                Response.ContentType = string.Empty;
                Response.AddHeader("content-disposition", "attachment; filename=test.pdf");
                Response.BinaryWrite(bytes); // create the file
                Response.Flush(); // send it to the client to download
            }
            catch (Exception e)
            {
                ViewBag.Error = JsonConvert.SerializeObject(e);
            }
            return View();
        }

        public ActionResult About()
        {
            try
            {
                var dir = HttpContext.Server.MapPath("~/");
                CreatePDF("test", dir);
            }
            catch (Exception e)
            {
                ViewBag.Error = JsonConvert.SerializeObject(e);
            }
            return View();
        }

        public ActionResult Contact()
        {
            try
            {
                var dir = HttpContext.Server.MapPath("~/");
                CreatePDF2("test2", dir);
            }
            catch (Exception e)
            {
                ViewBag.Error = JsonConvert.SerializeObject(e);
            }
            return View();
        }

        public void CreatePDF2(string fileName, string dir)
        {
            // Variables
            Warning[] warnings;
            string[] streamIds;
            string mimeType = string.Empty;
            string encoding = string.Empty;
            string extension = string.Empty;

            ReportViewer viewer = new ReportViewer();
            viewer.ProcessingMode = ProcessingMode.Local;
            viewer.LocalReport.ReportPath = dir + "rdlcReport\\test2.rdlc";
            byte[] bytes = viewer.LocalReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamIds, out warnings);

            // Now that you have all the bytes representing the PDF report, buffer it and send it to the client.
            Response.Buffer = true;
            Response.Clear();
            Response.ContentType = mimeType;
            Response.AddHeader("content-disposition", "attachment; filename=" + fileName + "." + extension);
            Response.BinaryWrite(bytes); // create the file
            Response.Flush(); // send it to the client to download
        }

        public void CreatePDF(string fileName, string dir)
        {
            // Variables
            Warning[] warnings;
            string[] streamIds;
            string mimeType = string.Empty;
            string encoding = string.Empty;
            string extension = string.Empty;


            // Setup the report viewer object and get the array of bytes
            ReportViewer viewer = new ReportViewer();
            viewer.ProcessingMode = ProcessingMode.Local;
            viewer.LocalReport.ReportPath = dir + "rdlcReport\\test.rdlc";

            ReportParameterCollection reportParameters = new ReportParameterCollection();
            reportParameters.Add(new ReportParameter("CompanyName", "Company 1234"));
            viewer.LocalReport.SetParameters(reportParameters);

            List<TestRender> testRender = new List<TestRender>();
            for (var i = 0; i < 100; i++)
            {
                testRender.Add(new TestRender { Col1 = $"this is value of {i} in table" });
            }

            DataTable dt = CreateDataTable<TestRender>(testRender);

            viewer.LocalReport.DataSources.Clear();
            viewer.LocalReport.DataSources.Add(new ReportDataSource("DataSet1", dt));

            byte[] bytes = viewer.LocalReport.Render("PDF", null, out mimeType, out encoding, out extension, out streamIds, out warnings);

            // Now that you have all the bytes representing the PDF report, buffer it and send it to the client.
            Response.Buffer = true;
            Response.Clear();
            Response.ContentType = mimeType;
            Response.AddHeader("content-disposition", "attachment; filename=" + fileName + "." + extension);
            Response.BinaryWrite(bytes); // create the file
            Response.Flush(); // send it to the client to download
        }

        public class TestRender
        {
            public string Col1 { get; set; }
        }

        public DataTable CreateDataTable<T>(IEnumerable<T> list)
        {
            Type type = typeof(T);
            var properties = type.GetProperties();

            DataTable dataTable = new DataTable();
            foreach (PropertyInfo info in properties)
            {
                dataTable.Columns.Add(new DataColumn(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType));
            }

            foreach (T entity in list)
            {
                object[] values = new object[properties.Length];
                for (int i = 0; i < properties.Length; i++)
                {
                    values[i] = properties[i].GetValue(entity);
                }

                dataTable.Rows.Add(values);
            }

            return dataTable;
        }


        public byte[] CovnertHTMLToPDF(string html, string css, byte[] font)
        {
            //Create a byte array that will eventually hold our final PDF
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
                        for (var i = 0; i < 1000; i++)
                        {
                            doc.Add(new Paragraph(String.Format("This is paragraph #{0}", i)));
                        }

                        doc.Close();
                    }
                }
                bytes = ms.ToArray();
            }

            //Read our sample PDF and apply page numbers
            using (var reader = new PdfReader(bytes))
            {
                using (var ms = new MemoryStream())
                {
                    using (var stamper = new PdfStamper(reader, ms))
                    {
                        int PageCount = reader.NumberOfPages;
                        for (int i = 1; i <= PageCount; i++)
                        {
                            RenderPdfFooter(stamper.GetOverContent(i), i);
                        }
                    }
                    bytes = ms.ToArray();
                }
            }
            return bytes;
        }

        public void RenderPdfFooter(PdfContentByte page, int PageNumber)
        {
            BaseFont basefont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false);
            iTextSharp.text.Font fontsize = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 20);
            page.SetColorFill(new BaseColor(51, 51, 51));
            var phare = new Phrase(PageNumber.ToString());
            phare.Font = fontsize;

            List<IElement> parsedText = ConvertToHtmlForColumnText("<table><tbody><tr><td>1</td><td>2</td><td>3</td></tr></tbody></table>");

            page.Rectangle(20, 20, 550, 75);
            page.Stroke();
            Rectangle rect = new Rectangle(20, 20, 550, 75);
            ColumnText ct = new ColumnText(page);
            ct.SetSimpleColumn(rect);
            foreach(var item in parsedText)
            {
                ct.AddElement(item);
            }
            ct.Go();

            ColumnText.ShowTextAligned(page, Element.ALIGN_RIGHT, phare, 575, 20, 0);
        }

        List<IElement> ConvertToHtmlForColumnText(String text)
        {
            ListElementHandler listHandler = new ListElementHandler();
            XMLWorkerHelper.GetInstance().ParseXHtml(listHandler, new StringReader(text));
            return listHandler.List;
        }
    }
}