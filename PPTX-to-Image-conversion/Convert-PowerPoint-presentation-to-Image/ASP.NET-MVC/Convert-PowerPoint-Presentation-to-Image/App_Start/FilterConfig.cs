using System.Web;
using System.Web.Mvc;

namespace Convert_PowerPoint_Presentation_to_Image
{
    public class FilterConfig
    {
        public static void RegisterGlobalFilters(GlobalFilterCollection filters)
        {
            filters.Add(new HandleErrorAttribute());
        }
    }
}
