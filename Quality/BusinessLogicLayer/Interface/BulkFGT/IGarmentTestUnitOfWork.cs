using System;
using ProductionDataAccessLayer.Interface;

namespace BusinessLogicLayer.Interface.BulkFGT
{
    public interface IGarmentTestUnitOfWork : IDisposable
    {
        IGarmentTestDetailShrinkageProvider ShrinkageProvider { get; }
        IGarmentDetailSpiralityProvider SpiralityProvider { get; }
        IGarmentTestDetailApperanceProvider AppearanceProvider { get; }
        IGarmentTestDetailFGPTProvider FgptProvider { get; }
        IGarmentTestDetailFGWTProvider FgwtProvider { get; }
        IGarmentTestDetailProvider DetailProvider { get; }

        void Commit();
        void Rollback();
    }
}
