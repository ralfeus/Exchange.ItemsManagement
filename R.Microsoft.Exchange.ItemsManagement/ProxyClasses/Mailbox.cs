using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses
{
    public class Mailbox
    {
        private readonly object instance;

        public string UserPrincipalName
        {
            get
            {
                var assembly = Loader.LoadAssemblyByPath(Helpers.GetExchangeBinariesDirectory() + "Microsoft.Exchange.Data.Directory.dll");
                return (string)instance.GetType().GetProperty("UserPrincipalName").GetValue(instance, new object[] {0});
            }
        }
    }
}
