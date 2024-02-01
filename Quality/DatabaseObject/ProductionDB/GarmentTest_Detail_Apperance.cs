using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    /*(GarmentTest_Detail_Apperance) 詳細敘述如下*/
    /// <summary>
    /// 
    /// </summary>
    /// <info>Author: Wei; Date: 2021/08/23 </info>
    /// <history>
    /// xx.  YYYY/MM/DD   Ver   Author      Comments
    /// ===  ==========  ====  ==========  ==========
    /// 01.  2021/08/23   1.00    Admin        Create
    /// </history>

    public class GarmentTest_Detail_Apperance
    {
        public long ID { get; set; }

        public int No { get; set; }

        private string _Type;

        public string Type
        {
            get => _Type ?? string.Empty;
            set => _Type = value;
        }

        private string _Wash1;

        public string Wash1
        {
            get => _Wash1 ?? string.Empty;
            set => _Wash1 = value;
        }

        private string _Wash2;

        public string Wash2
        {
            get => _Wash2 ?? string.Empty;
            set => _Wash2 = value;
        }

        private string _Wash3;

        public string Wash3
        {
            get => _Wash3 ?? string.Empty;
            set => _Wash3 = value;
        }

        private string _Comment;

        public string Comment
        {
            get => _Comment ?? string.Empty;
            set => _Comment = value;
        }

        public int Seq { get; set; }

        private string _Wash4;

        public string Wash4
        {
            get => _Wash4 ?? string.Empty;
            set => _Wash4 = value;
        }

        private string _Wash5;

        public string Wash5
        {
            get => _Wash5 ?? string.Empty;
            set => _Wash5 = value;
        }
    }

}
