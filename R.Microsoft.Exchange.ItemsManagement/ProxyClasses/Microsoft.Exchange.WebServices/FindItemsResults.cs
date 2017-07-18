using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class FindItemsResults<T> : BaseProxyClass, IEnumerable<Item> where T : Item
    {
        public FindItemsResults(object instance)
        {
            this.instance = instance;
        }

        internal bool MoreAvailable
        {
            get
            {
                return (bool)this.instanceType.GetProperty("MoreAvailable").GetValue(this.instance, null);
            }
        }

        internal Nullable<int> NextPageOffset
        {
            get
            {
                return (Nullable<int>)this.instance.GetType().GetProperty("NextPageOffset").GetValue(this.instance, null);
            }
        }

        internal int TotalCount
        {
            get
            {
                return (int)this.instanceType.GetProperty("TotalCount").GetValue(this.instance, null);
            }
        }

        public IEnumerator<Item> GetEnumerator()
        {
            return ((IEnumerable<Item>)this.instance).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable)this.instance).GetEnumerator();
        }
    }
}
