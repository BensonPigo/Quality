using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMailToProvider
    {
        IList<MailTo> Get(MailTo Item);

        IList<MailTo> GetCFTComments_ToAddress(RFT_OrderComments Item);
    }
}
