using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using DatabaseObject.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ManufacturingExecutionDataAccessLayer.Interface
{
    public interface IBulkFGTMailGroupProvider
    {
        IList<Quality_MailGroup> MailGroupGet(Quality_MailGroup quality_Mail);

        int MailGroupSave(Quality_MailGroup quality_Mail, int type);
    }
}
