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

        IList<GarmentDefectCode> GetGarmentDefectCode(GarmentDefectCode defectCode);

        IList<Area> GetArea(Area area);

        IList<DropDownList> GetDropDownList(DropDownList downList);
    }
}
