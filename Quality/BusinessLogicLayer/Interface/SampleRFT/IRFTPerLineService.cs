using DatabaseObject.ViewModel;

namespace BusinessLogicLayer.Interface
{
    public interface IRFTPerLineService
    {
        RFTPerLine_ViewModel GetQueryPara();

        RFTPerLine_ViewModel RFTPerLineQuery(string FactoryID, string Year, string Month);
    }
}
