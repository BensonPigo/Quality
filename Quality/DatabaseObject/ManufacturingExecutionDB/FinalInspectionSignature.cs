using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public class FinalInspectionSignature
    {
        public long FinalInspectionSignatureUkey { get; set; }
        public long Ukey { get; set; }
        public string FinalInspectionID { get; set; }
        public string JobTitle { get; set; }
        public string UserID { get; set; }
        public byte[] Signature { get; set; }
        public string AddName { get; set; }
        public DateTime AddDate { get; set; }
    }
}
