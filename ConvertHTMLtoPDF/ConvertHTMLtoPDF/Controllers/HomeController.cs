using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using T2P.Export_HTML_To_PDF;
using T2P.Export_HTML_To_PDF.Models;

namespace ConvertHTMLtoPDF.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Export()
        {
            try
            {
                PDFTemplate template = new PDFTemplate();
                template.url_css = HttpContext.Server.MapPath("~\\Template\\template.css");
                template.url_html = HttpContext.Server.MapPath("~\\Template\\template.html");
                template.url_html_details = HttpContext.Server.MapPath("~\\Template\\template_details.html");
                template.url_html_footer = HttpContext.Server.MapPath("~\\Template\\template_footer.html");
                template.dir = HttpContext.Server.MapPath("~\\Template");
                template.out_FileName = "testmvnas.pdf";
                ExportHelpers.ToPDF(template);
            }
            catch (Exception e)
            {
                ViewBag.Error = JsonConvert.SerializeObject(e);
            }
            return View();
        }

    }
}
