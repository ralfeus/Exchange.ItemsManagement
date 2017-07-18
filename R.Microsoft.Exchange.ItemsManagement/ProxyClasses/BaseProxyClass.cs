using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses
{
    class BaseProxyClass
    {
        private object _instance;
        protected Type instanceType { get; private set; }

        internal object instance
        {
            get { return this._instance; }
            set
            {
                this._instance = value;
                this.instanceType = value.GetType();
            }
        }
    }
}
