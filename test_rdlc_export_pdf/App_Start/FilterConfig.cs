using System.Web;
using System.Web.Mvc;

namespace test_rdlc_export_pdf
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
