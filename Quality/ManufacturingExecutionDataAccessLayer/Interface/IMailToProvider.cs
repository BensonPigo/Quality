using DatabaseObject.ManufacturingExecutionDB;
using System.Collections.Generic;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IMailToProvider
    {
        IList<MailTo> Get(MailTo Item);
    }
}
