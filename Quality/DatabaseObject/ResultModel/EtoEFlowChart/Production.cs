using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ResultModel.EtoEFlowChart
{
    public class Production
    {
        public decimal SubprocessRFT { get; set; }
        public string TestResult { get; set; }
        public decimal InlineRFT { get; set; }
        public decimal EndlineWFT { get; set; }
        public decimal MDPassRate { get; set; }

        public List<Production_SubprocessRFT> SubprocessRFT_Detail { get; set; }
        public Production_TestResult TestResult_Detail { get; set; }
        public List<Production_InlineRFT> InlineRFT_Detail { get; set; }
        public List<Production_EndlineWFT> EndlineWFT_Detail { get; set; }
        public List<Production_MDPassRate> MDPassRate_Detail { get; set; }
    }
    public class Production_SubprocessRFT
    {
        public string SubProcessID { get; set; }
        public decimal RFT { get; set; }
        public decimal SummaryRate { get; set; }
    }
    public class Production_TestResult
    {
        public string FGWTResult { get; set; }
        public string FGPTResult { get; set; }
    }
    public class Production_InlineRFT
    {
        public DateTime FirstInspDate { get; set; }
        public string FactoryID { get; set; }
        public string POID { get; set; }
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string Destination { get; set; }
        public string ProductType { get; set; }
        public string QCName { get; set; }
        public int CallInspectedQty { get; set; }
        public int RejectQty { get; set; }
        public decimal RFTRate { get; set; }
    }
    public class Production_EndlineWFT
    {
        public DateTime FirstInspDate { get; set; }
        public string FactoryID { get; set; }
        public string POID { get; set; }
        public string OrderID { get; set; }
        public string Article { get; set; }
        public string Destination { get; set; }
        public string ProductType { get; set; }
        public string QCName { get; set; }
        public int CallInspectedQty { get; set; }
        public int RejectQty { get; set; }
        public decimal WFTRate { get; set; }
    }
    public class Production_MDPassRate
    {
        public string OrderID { get; set; }
        public string Article { get; set; }
        public DateTime DeliveryDate { get; set; }
        public int MDFailQty { get; set; }
    }
}
