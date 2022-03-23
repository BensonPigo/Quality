using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionDataAccessLayer.Interface
{
    public interface IInspectionTypeProvider
    {
        IList<InspectionType> Get_InspectionType(string Function, string Category, string BrandID = "ADIDAS");
    }
}
