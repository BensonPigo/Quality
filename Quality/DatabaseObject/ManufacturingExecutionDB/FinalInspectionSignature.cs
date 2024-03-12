using DatabaseObject.Public;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DatabaseObject.ManufacturingExecutionDB
{
    public class FinalInspectionSignature : CompareBase
    {
        public bool Selected { get; set; }
        public string FinalInspectionID { get; set; }
        public string UserID { get; set; }
        public string UserName { get; set; }

        public string JobTitle { get; set; }
        public string Signature_Base64
        {
            get
            {
                byte[] tmp;
                if (this.Signature == null)
                {
                    tmp = new byte[1];
                }
                else
                {
                    tmp = this.Signature;
                }
                string base64 = "data:image/png;base64," + Convert.ToBase64String(tmp);
                return base64;
            }

            set
            {

            }
        }
        public byte[] Signature { get; set; }
        public string AddName { get; set; }
        public DateTime AddDate { get; set; }
    }
}
