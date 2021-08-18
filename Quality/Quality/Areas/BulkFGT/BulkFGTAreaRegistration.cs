using System.Web.Mvc;

namespace Quality.Areas.BulkFGT
{
    public class BulkFGTAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "BulkFGT";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "BulkFGT_default",
                "BulkFGT/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}