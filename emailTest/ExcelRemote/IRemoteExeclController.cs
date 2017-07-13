using Anko.Updater;

namespace Anko.ExcelRemote
{
    internal interface IRemoteExeclController
    {
        void RunExcelInit(OrdersParser parser, IDataUpdater dataUpdater);
    }
}