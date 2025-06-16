using System.Collections.Generic;
using System.Web.Mvc;
using FactoryDashBoardWeb.Helper;

namespace Quality.Helper
{
    public class DropdownLists
    {
        public List<SelectListItem> TemperatureList { get; set; }
        public List<SelectListItem> MachineList { get; set; }
        public List<SelectListItem> NeckList { get; set; }
        public List<SelectListItem> WashList { get; set; }
        public List<SelectListItem> TestResultPassList { get; set; }
        public List<SelectListItem> TestResultmmList { get; set; }
        public List<SelectListItem> MetalContentList { get; set; }
        public List<SelectListItem> ScaleList { get; set; }
    }

    public static class DropdownListHelper
    {
        public static DropdownLists Build(
            List<object> temperatures,
            List<string> machines,
            Dictionary<string, object> necks,
            List<string> washs,
            List<string> testResultPass,
            List<string> testResultmm,
            List<string> metalContents,
            List<string> scales)
        {
            var helper = new SetListItem();
            return new DropdownLists
            {
                TemperatureList = helper.ItemListBinding(temperatures),
                MachineList = helper.ItemListBinding(machines),
                NeckList = helper.ItemListBinding(necks),
                WashList = helper.ItemListBinding(washs),
                TestResultPassList = helper.ItemListBinding(testResultPass),
                TestResultmmList = helper.ItemListBinding(testResultmm),
                MetalContentList = helper.ItemListBinding(metalContents),
                ScaleList = helper.ItemListBinding(scales)
            };
        }
    }
}
