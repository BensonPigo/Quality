using System;
using ADOHelper.Utility;
using BusinessLogicLayer;
using BusinessLogicLayer.Interface.BulkFGT;
using ProductionDataAccessLayer.Interface;
using ProductionDataAccessLayer.Provider.MSSQL;

namespace BusinessLogicLayer.Service.BulkFGT
{
    public class GarmentTestUnitOfWork : IGarmentTestUnitOfWork
    {
        private readonly SQLDataTransaction _transaction;

        public IGarmentTestDetailShrinkageProvider ShrinkageProvider { get; }
        public IGarmentDetailSpiralityProvider SpiralityProvider { get; }
        public IGarmentTestDetailApperanceProvider AppearanceProvider { get; }
        public IGarmentTestDetailFGPTProvider FgptProvider { get; }
        public IGarmentTestDetailFGWTProvider FgwtProvider { get; }
        public IGarmentTestDetailProvider DetailProvider { get; }

        public GarmentTestUnitOfWork() : this(Common.ProductionDataAccessLayer)
        {
        }

        public GarmentTestUnitOfWork(string connection)
        {
            _transaction = new SQLDataTransaction(connection);
            ShrinkageProvider = new GarmentTestDetailShrinkageProvider(_transaction);
            SpiralityProvider = new GarmentDetailSpiralityProvider(_transaction);
            AppearanceProvider = new GarmentTestDetailApperanceProvider(_transaction);
            FgptProvider = new GarmentTestDetailFGPTProvider(_transaction);
            FgwtProvider = new GarmentTestDetailFGWTProvider(_transaction);
            DetailProvider = new GarmentTestDetailProvider(_transaction);
        }

        public void Commit() => _transaction.Commit();
        public void Rollback() => _transaction.RollBack();
        public void Dispose() => _transaction.CloseConnection();
    }
}
