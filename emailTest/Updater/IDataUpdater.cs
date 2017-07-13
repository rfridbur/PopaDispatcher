using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Anko.Updater
{
    interface IDataUpdater
    {
        void updateCustomerList(IList<Common.Customer> list);
        void updateAgentList(IList<Common.Agent> list);
        void updateShippingCompanyList(IList<Common.ShippingCompany> list);
        void updateOrderList(IList<Common.Order> list);
    }
}
