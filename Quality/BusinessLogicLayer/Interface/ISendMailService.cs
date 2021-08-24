using DatabaseObject.RequestModel;
using DatabaseObject.ResultModel;

namespace BusinessLogicLayer.Interface
{
    public interface ISendMailService
    {
        SendMail_Result SendMail(SendMail_Request SendMail_Request);
    }
}
