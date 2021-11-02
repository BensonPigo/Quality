using DatabaseObject.ProductionDB;
using System;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleArticleProvider
    {
        IList<Style_Article> Get(Style_Article Item);
    }
}
