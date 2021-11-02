using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ManufacturingExecutionDB
{
    public class Quality_MailGroup
    {
        /// <summary>廠代</summary>
        [Display(Name = "廠代")]
        public string FactoryID { get; set; }

        /// <summary>分類代碼</summary>
        [Display(Name = "分類代碼")]
        public string Type { get; set; }

        /// <summary>群組名稱</summary>
        [Display(Name = "群組名稱")]
        public string GroupName { get; set; }

        /// <summary>收件人</summary>
        [Display(Name = "收件人")]
        public string ToAddress { get; set; }

        /// <summary>副本</summary>
        [Display(Name = "副本")]
        public string CcAddress { get; set; }
    }
}
