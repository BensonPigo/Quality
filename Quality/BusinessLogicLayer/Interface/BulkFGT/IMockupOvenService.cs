using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupOvenService
    {
        MockupOven_ViewModel GetMockupOven(MockupOven_Request MockupOven);

        List<SelectListItem> GetAccessoryRefNo(AccessoryRefNo_Request Request);

        List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request);

        MockupOven_ViewModel Create(MockupOven_ViewModel MockupOven);

        MockupOven_ViewModel Update(MockupOven_ViewModel MockupOven);

        MockupOven_ViewModel Delete(MockupOven_ViewModel MockupOven);

        MockupOven_ViewModel GetPDF(MockupOven_ViewModel MockupOven, bool test = false);
    }
}
