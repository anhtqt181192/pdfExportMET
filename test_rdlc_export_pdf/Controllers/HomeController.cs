using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Reporting.WebForms;
using Microsoft.ReportingServices.Diagnostics.Internal;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Mvc;

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
            string html = System.IO.File.ReadAllText(HttpContext.Server.MapPath("~\\Template\\template.html"));
            string css = System.IO.File.ReadAllText(HttpContext.Server.MapPath("~\\Template\\template.css"));
            var bytes = CovnertHTMLToPDF(html, css);
            ByteArrayToFile(bytes, HttpContext.Server.MapPath("~\\Template\\testfile.pdf"));
            return View();
        }

        public void ByteArrayToFile(byte[] byteArray, string fileName)
        {
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(byteArray, 0, byteArray.Length);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in process: {0}", ex);
            }
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


        public byte[] CovnertHTMLToPDF(string html, string css)
        {
            //Create a byte array that will eventually hold our final PDF
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(PageSize.A4, 25, 25, 25, 25))
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
            var testFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "test.pdf");
            System.IO.File.WriteAllBytes(testFile, bytes);

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
                            ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_RIGHT, new Phrase(String.Format("{0}", i)), 585, 10, 0);
                        }
                    }
                    bytes = ms.ToArray();
                }
            }

            return bytes;
        }
    }
}