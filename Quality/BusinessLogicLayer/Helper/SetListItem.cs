using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;

namespace BusinessLogicLayer.Helper
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

        public List<SelectListItem> ItemListBinding(List<object> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.ToString(),
                    Value = item.ToString()
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, object> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Key.ToString(),
                    Value = item.Value.ToString()
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, int> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Key.ToString(),
                    Value = item.Value.ToString()
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }

        public List<SelectListItem> ItemListBinding(Dictionary<string, string> Options)
        {
            List<SelectListItem> result_itemList = new List<SelectListItem>();
            foreach (var item in Options)
            {
                SelectListItem i = new SelectListItem()
                {
                    Text = item.Key.ToString(),
                    Value = item.Value.ToString()
                };
                result_itemList.Add(i);
            }

            return result_itemList;
        }
    }
}
