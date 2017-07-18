using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class ItemView : BaseProxyClass
    {
        internal ItemView(int maxItems)
        {
            this.instance = ClassFactory.CreateInstance("ItemView", new object[] {maxItems}, "Microsoft_Exchange_WebServices", true);
        }

        internal int Offset
        {
            set
            {
                instance.GetType().GetProperty("Offset").SetValue(instance, value, null);
            }
        }
    }
}
