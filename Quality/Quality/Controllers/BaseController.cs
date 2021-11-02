using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Quality.Controllers
{
    public class BaseController : Controller
    {
        #region 參數屬性區
        public string UserID
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["UserID"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["UserID"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["UserID"] = value; }
        }

        public string BulkFGT_Brand
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["BulkFGT_Brand"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["BulkFGT_Brand"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["BulkFGT_Brand"] = value; }
        }

        public string UserMail
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["UserMail"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["UserMail"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["UserMail"] = value; }
        }

        public string MDivisionID
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["MDivisionID"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["MDivisionID"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["MDivisionID"] = value; }
        }


        public string FactoryID
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["FactoryID"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["FactoryID"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["FactoryID"] = value; }
        }

        public string Line
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["Line"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["Line"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["Line"] = value; }
        }
        public string Brand
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["Brand"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["Brand"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["Brand"] = value; }
        }

        public DateTime WorkDate
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["WorkDate"] == null) { return new DateTime(); }
                else { return (DateTime)System.Web.HttpContext.Current.Session["WorkDate"]; }
            }
            set { System.Web.HttpContext.Current.Session["WorkDate"] = value; }
        }

        public List<Quality_Menu> MenuList
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["MenuList"] == null) { return new List<Quality_Menu>(); }
                else { return (List<Quality_Menu>)System.Web.HttpContext.Current.Session["MenuList"]; }
            }
            set { System.Web.HttpContext.Current.Session["MenuList"] = value; }
        }

        public string SelectedMenu
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["SelectedMenu"] == null) { return ""; }
                else { return System.Web.HttpContext.Current.Session["SelectedMenu"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["SelectedMenu"] = value; }
        }

        public string OnlineHelp
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["OnlineHelp"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["OnlineHelp"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["OnlineHelp"] = value; }
        }

        public List<string> Factorys
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["Factorys"] == null) { return new List<string>(); }
                else { return (List<string>)System.Web.HttpContext.Current.Session["Factorys"]; }
            }
            set { System.Web.HttpContext.Current.Session["Factorys"] = value; }
        }

        public List<string> Lines
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["Lines"] == null) { return new List<string>(); }
                else { return (List<string>)System.Web.HttpContext.Current.Session["Lines"]; }
            }
            set { System.Web.HttpContext.Current.Session["Lines"] = value; }
        }
        public List<string> Brands
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["Brands"] == null) { return new List<string>(); }
                else { return (List<string>)System.Web.HttpContext.Current.Session["Brands"]; }
            }
            set { System.Web.HttpContext.Current.Session["Brands"] = value; }
        }

        public List<Inspection_ViewModel> SelectItemData
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["SelectItemData"] == null) { return new List<Inspection_ViewModel>(); }
                else { return (List<Inspection_ViewModel>)System.Web.HttpContext.Current.Session["SelectItemData"]; }
            }
            set { System.Web.HttpContext.Current.Session["SelectItemData"] = value; }
        }

        public string TargetArea
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["TargetArea"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["TargetArea"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["TargetArea"] = value; }
        }

        public string TargetController
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["TargetController"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["TargetController"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["TargetController"] = value; }
        }

        public string TargetAction
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["TargetAction"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["TargetAction"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["TargetAction"] = value; }
        }


        public string TargetPKey_Parameter
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["TargetPKey_Parameter"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["TargetPKey_Parameter"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["TargetPKey_Parameter"] = value; }
        }

        public string TargetPKey_Value
        {
            get
            {
                if (System.Web.HttpContext.Current.Session["TargetPKey_Value"] == null) { return string.Empty; }
                else { return System.Web.HttpContext.Current.Session["TargetPKey_Value"].ToString(); }
            }
            set { System.Web.HttpContext.Current.Session["TargetPKey_Value"] = value; }
        }


        #endregion

        // 檢查是否登入
        public bool CheckSession()
        {
            if (Session == null || Session.Keys.Count == 1)
            {
                if (Request.IsAjaxRequest())//是Ajax的話
                {
                    Response.StatusCode = 440;//Login timeout
                    Response.End();
                }
                else
                {
                    Response.Write("Please login again。</br></br>");
                    Response.End();
                }       
                return false;
            }
            return true;
        }
    }
}