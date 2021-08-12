using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ProductionDB;
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

        List<RFT_OrderComments_ViewModel> RFT_OrderCommentsGet(RFT_OrderComments rFT_OrderComments);

        RFT_OrderComments_ViewModel RFT_OrderCommentsSave(List<RFT_OrderComments> rFT_OrderComments);

        RFT_OrderComments_ViewModel RFT_OrderCommentsSendMail(RFT_OrderComments rFT_OrderComments);

        RFT_PicDuringDummyFitting RFT_PicDuringDummyFittingGet(RFT_PicDuringDummyFitting picDuringDummyFitting);

        RFT_PicDuringDummyFitting_ViewModel RFT_PicDuringDummyFittingSave(RFT_PicDuringDummyFitting picDuringDummyFitting);
    }
}
