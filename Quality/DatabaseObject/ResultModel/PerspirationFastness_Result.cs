using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel
{
    public class PerspirationFastness_Result : BaseResult
    {
        public PerspirationFastness_Main Main { get; set; }
        public List<PerspirationFastness_Detail> Details { get; set; } = new List<PerspirationFastness_Detail>();
    }

    public class PerspirationFastness_Main
    {
        public string POID { get; set; }
        public string StyleID { get; set; }
        public string BrandID { get; set; }
        public string SeasonID { get; set; }
        public DateTime? CutInline { get; set; }
        public DateTime? MinSciDelivery { get; set; }
        public DateTime? TargetLeadTime { get; set; }
        public DateTime? CompletionDate { get; set; }
        public decimal LabPerspirationFastnessPercent { get; set; }
        public string Remark { get; set; }
        public string CreateBy { get; set; }
        public string EditBy { get; set; }
    }

    public class PerspirationFastness_Detail
    {
        public string TestNo { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Result { get; set; }
        public string Inspector { get; set; }
        public string Remark { get; set; }
        public string LastUpdate { get; set; }
        public string Status { get; set; }
    }


    public class PerspirationFastness_Detail_Result : BaseResult
    {
        public List<string> ScaleIDs { get; set; }
        public PerspirationFastness_Detail_Main Main { get; set; } = new PerspirationFastness_Detail_Main();
        public List<PerspirationFastness_Detail_Detail> Details { get; set; } = new List<PerspirationFastness_Detail_Detail>();
    }

    public class PerspirationFastness_Detail_Main
    {
        public long ID { get; set; }
        public string TestNo { get; set; }
        public string POID { get; set; }
        public DateTime? InspDate { get; set; }
        public string Article { get; set; }
        public string Inspector { get; set; }
        public string InspectorName { get; set; }
        public string Result { get; set; }
        public string Remark { get; set; }
        public string Status { get; set; }
        public string MetalContent { get; set; }
        public string Temperature { get; set; }
        public string Time { get; set; }
        public byte[] TestBeforePicture { get; set; }
        public byte[] TestAfterPicture { get; set; }
    }

    public class PerspirationFastness_Detail_Detail : CompareBase
    {
        public DateTime? SubmitDate { get; set; }
        public string PerspirationFastnessGroup { get; set; }
        public string SEQ { get; set; }
        public string Roll { get; set; }
        public string Dyelot { get; set; }
        public string Refno { get; set; }
        public string SCIRefno { get; set; }
        public string ColorID { get; set; }
        public string Result { get; set; }


        public string AlkalineChangeScale { get; set; }
        public string AlkalineResultChange { get; set; }

        public string AlkalineAcetateScale { get; set; }
        public string AlkalineResultAcetate { get; set; }

        public string AlkalineCottonScale { get; set; }
        public string AlkalineResultCotton { get; set; }

        public string AlkalineNylonScale { get; set; }
        public string AlkalineResultNylon { get; set; }

        public string AlkalinePolyesterScale { get; set; }
        public string AlkalineResultPolyester { get; set; }

        public string AlkalineAcrylicScale { get; set; }
        public string AlkalineResultAcrylic { get; set; }

        public string AlkalineWoolScale { get; set; }
        public string AlkalineResultWool { get; set; }

        public string AcidChangeScale { get; set; }
        public string AcidResultChange { get; set; }

        public string AcidAcetateScale { get; set; }
        public string AcidResultAcetate { get; set; }

        public string AcidCottonScale { get; set; }
        public string AcidResultCotton { get; set; }

        public string AcidNylonScale { get; set; }
        public string AcidResultNylon { get; set; }

        public string AcidPolyesterScale { get; set; }
        public string AcidResultPolyester { get; set; }

        public string AcidAcrylicScale { get; set; }
        public string AcidResultAcrylic { get; set; }

        public string AcidWoolScale { get; set; }
        public string AcidResultWool { get; set; }        

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

    public class PerspirationFastness_Excel
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

        public string MetalContent { get; set; }
        public string AlkalineChangeScale { get; set; }
        public string AlkalineResultChange { get; set; }

        public string AlkalineAcetateScale { get; set; }
        public string AlkalineResultAcetate { get; set; }

        public string AlkalineCottonScale { get; set; }
        public string AlkalineResultCotton { get; set; }

        public string AlkalineNylonScale { get; set; }
        public string AlkalineResultNylon { get; set; }

        public string AlkalinePolyesterScale { get; set; }
        public string AlkalineResultPolyester { get; set; }

        public string AlkalineAcrylicScale { get; set; }
        public string AlkalineResultAcrylic { get; set; }

        public string AlkalineWoolScale { get; set; }
        public string AlkalineResultWool { get; set; }

        public string AcidChangeScale { get; set; }
        public string AcidResultChange { get; set; }

        public string AcidAcetateScale { get; set; }
        public string AcidResultAcetate { get; set; }

        public string AcidCottonScale { get; set; }
        public string AcidResultCotton { get; set; }

        public string AcidNylonScale { get; set; }
        public string AcidResultNylon { get; set; }

        public string AcidPolyesterScale { get; set; }
        public string AcidResultPolyester { get; set; }

        public string AcidAcrylicScale { get; set; }
        public string AcidResultAcrylic { get; set; }

        public string AcidWoolScale { get; set; }
        public string AcidResultWool { get; set; }
        public string Remark { get; set; }
        public string Inspector { get; set; }

        public Byte[] TestBeforePicture { get; set; }

        public Byte[] TestAfterPicture { get; set; }

    }
}
