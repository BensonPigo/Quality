using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace FactoryDashBoardWeb.Helper
{
    public class SetListItem
    {
        public List<SelectListItem> ItemListBinding(List<string> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item,
                    Value = item
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }
    }
}