using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_FGPT) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>


    public class GarmentTest_Detail_FGPT
    {
        public long ID { get; set; }

        public int No { get; set; }

        private string _Location;

        public string Location
        {
            get => _Location ?? string.Empty;
            set => _Location = value;
        }

        private string _Type;

        public string Type
        {
            get => _Type ?? string.Empty;
            set => _Type = value;
        }

        private string _TestName;

        public string TestName
        {
            get => _TestName ?? string.Empty;
            set => _TestName = value;
        }

        private string _TestDetail;

        public string TestDetail
        {
            get => _TestDetail ?? string.Empty;
            set => _TestDetail = value;
        }

        public int Criteria { get; set; }

        private string _TestResult;

        public string TestResult
        {
            get => _TestResult ?? string.Empty;
            set => _TestResult = value;
        }

        private string _TestUnit;

        public string TestUnit
        {
            get => _TestUnit ?? string.Empty;
            set => _TestUnit = value;
        }

        public string StandardRemark { get; set; }

        private string _Wash5;

        public string Wash5
        {
            get => _Wash5 ?? string.Empty;
            set => _Wash5 = value;
        }

        public int Seq { get; set; }

        public int TypeSelection_VersionID { get; set; }

        public int TypeSelection_Seq { get; set; }

        public bool IsOriginal { get; set; }
    }

}
