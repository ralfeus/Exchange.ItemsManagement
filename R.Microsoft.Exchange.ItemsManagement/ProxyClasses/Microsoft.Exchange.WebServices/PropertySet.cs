using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class PropertySet : BaseProxyClass
    {
        internal PropertySet(BasePropertySet basePropertySet)
        {
            this.instance = ClassFactory.CreateInstance("PropertySet", new object[] { basePropertySet }, "Microsoft_Exchange_WebServices", true);
        }

        internal BodyType RequestedBodyType
        {
            set
            {
                this.instance.GetType().GetProperty("RequestedBodyType").SetValue(this.instance, value, null);
            }
        }
    }
}
