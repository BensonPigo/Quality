using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface.SampleRFT
{
    public interface IPicturesDummyService
    {
        RFT_PicDuringDummyFitting_ViewModel Get_PicturesDummy_Result(RFT_PicDuringDummyFitting_ViewModel Req);
        RFT_PicDuringDummyFitting_ViewModel Check_OrderID_Exists(string orderID);
    }
}
