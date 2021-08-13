using DatabaseObject.ManufacturingExecutionDB;
using DatabaseObject.ResultModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogicLayer.Interface
{
    public interface IBulkFGTMailGroup_Service
    {
        List<Quality_MailGroup> MailGroupGet(Quality_MailGroup quality_Mail);

        Quality_MailGroup_ResultModel MailGroupSave(Quality_MailGroup quality_Mail);
    }
}
