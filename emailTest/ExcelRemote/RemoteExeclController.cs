using Anko.Updater;
using System;

namespace Anko.ExcelRemote
{
    class RemoteExeclController : MarshalByRefObject, IRemoteExeclController
    {
        public static IDataUpdater DataUpdater;

        public void RunExcelInit(OrdersParser parser, IDataUpdater dataUpdater)
        {
            DataUpdater = dataUpdater;

            OrdersParser._Form = parser;

            Excel.init();
            Outlook.init();

            // create temp results folder
            Utils.createResultsFolder();

            // parse local DB
            Excel.getDetailsFromLocalDb();

            // fetch and save to file the most updated orders excel file
            Outlook.readLastOrdersFile();

            // parse the orders DB
            Excel.getOrderDetails();
        }
    }
}
