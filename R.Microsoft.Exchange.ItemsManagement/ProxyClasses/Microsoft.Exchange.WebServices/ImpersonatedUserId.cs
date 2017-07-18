using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class ImpersonatedUserId : BaseProxyClass
    {
        internal ImpersonatedUserId(ConnectingIdType connectingIdType, string id)
        {
            this.instance = ClassFactory.CreateInstance(
                "ImpersonatedUserId", new object[] {connectingIdType, id}, "Microsoft_Exchange_WebServices", true);
        }

    }
}
