using DatabaseObject.RequestModel;
using DatabaseObject.ViewModel.BulkFGT;
using System.Collections.Generic;
using System.Web.Mvc;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IMockupCrockingService
    {
        MockupCrocking_ViewModel GetMockupCrocking(MockupCrocking_Request MockupCrocking);

        List<SelectListItem> GetArtworkTypeID(StyleArtwork_Request Request);

        MockupCrocking_ViewModel Create(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel Update(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel Delete(MockupCrocking_ViewModel MockupCrocking);

        MockupCrocking_ViewModel GetPDF(MockupCrocking_ViewModel MockupCrocking, bool test = false);
    }
}
