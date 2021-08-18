using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel;
using System.Collections.Generic;
using static BusinessLogicLayer.Service.InspectionService;

namespace BusinessLogicLayer.Interface
{
    public interface IInspectionService
    {
        IList<Inspection_ViewModel> GetSelectItemData(Inspection_ViewModel inspection_ViewModel);

        IList<Inspection_ViewModel> CheckSelectItemData(Inspection_ViewModel inspection_ViewModel, SelectType type);

        Inspection_ViewModel GetTop3(Inspection_ViewModel inspection_ViewModel);

        IList<GarmentDefectType> GetGarmentDefectType();

        IList<MailTo> GetMailTo(MailTo mailTo);

        IList<GarmentDefectCode> GetGarmentDefectCode(GarmentDefectCode defectCode);

        IList<Area> GetArea(Area area);

        IList<DropDownList> GetDropDownList(DropDownList downList);

        InspectionSave_ViewModel SaveRFTInspection(InspectionSave_ViewModel inspections);

        IList<ReworkCard> GetReworkCards(ReworkCard rework);

        List<ReworkList_ViewModel> GetReworkList(ReworkList_ViewModel reworkList);

        InspectionSave_ViewModel SaveReworkListAction(List<RFT_Inspection> rFT_Inspections, ReworkListType reworkListType);

        InspectionSave_ViewModel SaveReworkListAddReject(RFT_Inspection_Detail detail);

        InspectionSave_ViewModel SaveReworkListDelete(LogIn_Request logIn_Request, List<RFT_Inspection> rFT_Inspection);

        List<DQSReason> GetDQSReason(DQSReason dQSReason);

        List<RFT_Inspection_Measurement_ViewModel> GetMeasurement(string OrderID, string SizeCode, string UserID);

        RFT_Inspection_Measurement_ViewModel SaveMeasurement(List<RFT_Inspection_Measurement> Measurement);

        List<RFT_OrderComments_ViewModel> GetRFT_OrderComments(RFT_OrderComments rFT_OrderComments);

        RFT_OrderComments_ViewModel SaveRFT_OrderComments(List<RFT_OrderComments> rFT_OrderComments);

        RFT_OrderComments_ViewModel SendMailRFT_OrderComments(RFT_OrderComments rFT_OrderComments);

        RFT_PicDuringDummyFitting GetRFT_PicDuringDummyFitting(RFT_PicDuringDummyFitting picDuringDummyFitting);

        RFT_PicDuringDummyFitting_ViewModel SaveRFT_PicDuringDummyFitting(RFT_PicDuringDummyFitting picDuringDummyFitting);
    }
}
