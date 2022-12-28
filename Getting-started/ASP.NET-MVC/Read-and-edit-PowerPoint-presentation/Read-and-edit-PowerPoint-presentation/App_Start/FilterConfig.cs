using System.Web;
using System.Web.Mvc;

namespace Read_and_edit_PowerPoint_presentation
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
