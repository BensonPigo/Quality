using DatabaseObject.ProductionDB;
using DatabaseObject.RequestModel;
using System.Collections.Generic;

namespace ProductionDataAccessLayer.Interface
{
    public interface IStyleArtworkProvider
    {
        IList<Style_Artwork> GetArtworkTypeID(StyleArtwork_Request Item);
    }
}
