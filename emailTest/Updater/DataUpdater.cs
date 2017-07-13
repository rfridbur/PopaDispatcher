using System;
using System.Collections.Generic;

namespace Anko.Updater
{
    // class is responsible to copy data generated inside the APP domain into outside data structures
    // so in case APP domain is closed, the data is not lost
    class DataUpdater : MarshalByRefObject, IDataUpdater
    {
        // function updates agents list
        public void updateAgentList(IList<Common.Agent> list)
        {
            Common.agentList = new List<Common.Agent>(list);
        }

        // function updates customers list
        public void updateCustomerList(IList<Common.Customer> list)
        {
            Common.customerList = new List<Common.Customer>(list);
        }

        // function updates orders list
        public void updateOrderList(IList<Common.Order> list)
        {
            Common.orderList = new List<Common.Order>(list);
        }

        // function updates shipping company list
        public void updateShippingCompanyList(IList<Common.ShippingCompany> list)
        {
            Common.shippingCompanyList = new List<Common.ShippingCompany>(list);
        }
    }
}
