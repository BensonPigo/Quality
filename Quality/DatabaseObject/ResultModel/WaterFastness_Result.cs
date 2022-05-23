using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class WaterFastness_Result : BaseResult
    {
        public WaterFastness_Main Main { get; set; }
        public List<WaterFastness_Detail> Details { get; set; } = new List<WaterFastness_Detail>();
    }

    public class WaterFastness_Main
    {
        public string POID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? CutInline { get; set; }
        public DateTime? MinSciDelivery { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal LabWaterFastnessPercent { get; set; }
        public string Remark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }
    }

    public class WaterFastness_Detail
    {
        public string ReportNo { get; set; }
        public string TestNo { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Result { get; set; }
        public string Inspector { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
        public string Status { get; set; }
    }


    public class WaterFastness_Detail_Result : BaseResult
    {
        public string MDivisionID { get; set; }
        public List<string> ScaleIDs { get; set; }
        public WaterFastness_Detail_Main Main { get; set; } = new WaterFastness_Detail_Main();
        public List<WaterFastness_Detail_Detail> Details { get; set; } = new List<WaterFastness_Detail_Detail>();
    }

    public class WaterFastness_Detail_Main
    {
        public long ID { get; set; }
        public string ReportNo { get; set; }
        public string TestNo { get; set; }
        public string POID { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Inspector { get; set; }
        public string InspectorName { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string Status { get; set; }
        public string Temperature { get; set; }
        public string Time { get; set; }
        public byte[] TestBeforePicture { get; set; }
        public byte[] TestAfterPicture { get; set; }
    }

    public class WaterFastness_Detail_Detail : CompareBase
    {
        public DateTime? SubmitDate { get; set; }
        public string WaterFastnessGroup { get; set; }
        public string SEQ { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public string Refno { get; set; }
        public string SCIRefno { get; set; }
        public string ColorID { get; set; }
        public string Result { get; set; }
        public string ChangeScale { get; set; }
        public string ResultChange { get; set; }

        public string AcetateScale { get; set; }
        public string ResultAcetate { get; set; }

        public string CottonScale { get; set; }
        public string ResultCotton { get; set; }

        public string NylonScale { get; set; }
        public string ResultNylon { get; set; }

        public string PolyesterScale { get; set; }
        public string ResultPolyester { get; set; }

        public string AcrylicScale { get; set; }
        public string ResultAcrylic { get; set; }

        public string WoolScale { get; set; }
        public string ResultWool { get; set; }

        public string Remark { get; set; }
        public string LastUpdate { get; set; }
        public string Temperature { get; set; }
        public string Time { get; set; }

        public string Seq1
        {
            get
            {
                if (string.IsNullOrEmpty(this.SEQ))
                {
                    return string.Empty;
                }

                if (!this.SEQ.Contains("-"))
                {
                    return string.Empty;
                }

                return this.SEQ.Split('-')[0].Trim();
            }
        }

        public string Seq2
        {
            get
            {
                if (string.IsNullOrEmpty(this.SEQ))
                {
                    return string.Empty;
                }

                if (!this.SEQ.Contains("-"))
                {
                    return string.Empty;
                }

                return this.SEQ.Split('-')[1].Trim();
            }
        }
    }

    public class WaterFastness_Excel
    {
        public DateTime? SubmitDate { get; set; }
        public string SeasonID { get; set; }
        public string BrandID { get; set; }
        public string StyleID { get; set; }
        public string POID { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public int Temperature { get; set; }
        public int Time { get; set; }
        public string SCIRefno_Color { get; set; }
        public string ChangeScale { get; set; }
        public string ResultChange { get; set; }
        public string AcetateScale { get; set; }
        public string ResultAcetate { get; set; }
        public string CottonScale { get; set; }
        public string ResultCotton { get; set; }
        public string NylonScale { get; set; }
        public string ResultNylon { get; set; }
        public string PolyesterScale { get; set; }
        public string ResultPolyester { get; set; }
        public string AcrylicScale { get; set; }
        public string ResultAcrylic { get; set; }
        public string WoolScale { get; set; }
        public string ResultWool { get; set; }
        public string Remark { get; set; }
        public string Inspector { get; set; }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }

    }
}
