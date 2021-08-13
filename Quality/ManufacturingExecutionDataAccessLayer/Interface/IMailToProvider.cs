using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMailToProvider
    {
        IList<MailTo> Get(MailTo Item);

        IList<MailTo> GetMR_SMR_MailAddress(RFT_OrderComments Item, string MailID);
    }
}
