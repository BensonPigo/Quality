using System.Collections.Generic;
using System.Web.Mvc;

namespace FactoryDashBoardWeb.Helper
{
    /// <summary>
    /// Compatibility wrapper that forwards to <see cref="SetListItem"/>
    /// </summary>
    public class DropdownLists
    {
        private readonly SetListItem _helper = new SetListItem();

        public List<SelectListItem> ItemListBinding(List<string> options)
        {
            return _helper.ItemListBinding(options);
        }

        public List<SelectListItem> ItemListBinding(List<object> options)
        {
            return _helper.ItemListBinding(options);
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, object> options)
        {
            return _helper.ItemListBinding(options);
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, int> options)
        {
            return _helper.ItemListBinding(options);
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, string> options)
        {
            return _helper.ItemListBinding(options);
        }
    }
}
