using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T2P.Export_HTML_To_PDF.Models
{
    public class PDFTemplate
    {
        public string url_html { get; set; }
        public string url_html_details { get; set; }
        public string url_html_footer { get; set; }
        public string url_css { get; set; }
        public string url_font { get; set; }
        public string dir { get; set; }
        public string out_FileName { get; set; }
    }
}
