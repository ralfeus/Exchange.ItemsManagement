using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class FindFoldersResults : BaseProxyClass, IEnumerable<Folder>
    {
        public FindFoldersResults(object instance)
        {
            this.instance = instance;
        }

        public IEnumerator GetEnumerator()
        {
            return ((IEnumerable)this.instance).GetEnumerator();
        }

        IEnumerator<Folder> IEnumerable<Folder>.GetEnumerator()
        {
            return ((IEnumerable<Folder>)this.instance).GetEnumerator();
        }
    }
}
