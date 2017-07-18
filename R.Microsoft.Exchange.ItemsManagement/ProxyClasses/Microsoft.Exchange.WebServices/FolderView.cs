using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class FolderView : BaseProxyClass
    {
        internal FolderView(int maxItems)
        {
            this.instance = ClassFactory.CreateInstance("FolderView", new object[] {maxItems}, "Microsoft_Exchange_WebServices", true);
        }

        internal FolderTraversal Traversal
        {
            set
            {
                instance.GetType().GetProperty("Traversal").SetValue(instance, value, null);
            }
        }
    }
}
