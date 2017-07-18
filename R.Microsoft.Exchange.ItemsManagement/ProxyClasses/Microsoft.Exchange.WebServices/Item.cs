using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class Item : BaseProxyClass
    {
        internal Item()
        {
            this.instance = ClassFactory.CreateInstance("Item", "Microsoft_Exchange_WebServices", true);
        }

        internal void Load(PropertySet propertySet)
        {
            this.instanceType.GetMethod("Load").Invoke(this.instance, new object[] { propertySet.instance });
        }
    }
}
