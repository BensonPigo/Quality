using System;
using System.ComponentModel.DataAnnotations;
namespace DatabaseObject.ProductionDB
{
    public class ColorFastness2
    {
        public string ID { get; set; }

        public string ReportNo { get; set; }
        public string POID { get; set; }

        public decimal TestNo { get; set; }

        public DateTime? InspDate { get; set; }

        public string Article { get; set; }

        public string Result { get; set; }

        public string Status { get; set; }

        public string Inspector { get; set; }

        public string Remark { get; set; }

        public string addName { get; set; }

        public DateTime? addDate { get; set; }

        public string EditName { get; set; }

        public DateTime? EditDate { get; set; }

        public int Temperature { get; set; }
        public string BrandID { get; set; }

        public int Cycle { get; set; }
        public int CycleTime { get; set; }

        private string _Detergent;
        
        public string Detergent
        {

            get
            {
                if (_Detergent == null)
                {
                    _Detergent = string.Empty;
                }

                return _Detergent;
            }

            set
            {
                _Detergent = value;
            }
        }


        private string _Machine;

        public string Machine
        {
            get
            {
                if (_Machine == null)
                {
                    _Machine = string.Empty;
                }

                return _Machine;
            }

            set
            {
                _Machine = value;
            }
        }
        public string Drying { get; set; }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }

    }

    public class ColorFastness
    {
        public string ID { get; set; }

        private string _ReportNo;

        public string ReportNo
        {
            get => _ReportNo ?? string.Empty;
            set => _ReportNo = value;
        }

        private string _POID;

        public string POID
        {
            get => _POID ?? string.Empty;
            set => _POID = value;
        }

        public decimal TestNo { get; set; }

        public DateTime? InspDate { get; set; }

        private string _Article;

        public string Article
        {
            get => _Article ?? string.Empty;
            set => _Article = value;
        }

        private string _Result;

        public string Result
        {
            get => _Result ?? string.Empty;
            set => _Result = value;
        }

        private string _Status;

        public string Status
        {
            get => _Status ?? string.Empty;
            set => _Status = value;
        }

        private string _Inspector;

        public string Inspector
        {
            get => _Inspector ?? string.Empty;
            set => _Inspector = value;
        }

        private string _Remark;

        public string Remark
        {
            get => _Remark ?? string.Empty;
            set => _Remark = value;
        }

        private string _addName;

        public string addName
        {
            get => _addName ?? string.Empty;
            set => _addName = value;
        }

        public DateTime? addDate { get; set; }

        private string _EditName;

        public string EditName
        {
            get => _EditName ?? string.Empty;
            set => _EditName = value;
        }

        public DateTime? EditDate { get; set; }

        public int Temperature { get; set; }

        private string _BrandID;

        public string BrandID
        {
            get => _BrandID ?? string.Empty;
            set => _BrandID = value;
        }

        public int Cycle { get; set; }

        public int CycleTime { get; set; }

        private string _Detergent;

        public string Detergent
        {
            get => _Detergent ?? string.Empty;
            set => _Detergent = value;
        }

        private string _Machine;

        public string Machine
        {
            get => _Machine ?? string.Empty;
            set => _Machine = value;
        }

        private string _Drying;

        public string Drying
        {
            get => _Drying ?? string.Empty;
            set => _Drying = value;
        }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }
        public string _Approver;


        public string Approver
        { 
            get => _Approver??string.Empty; 
            set => _Approver = value; 
        }
        public string ApproverName { get; set; }

    }

}
