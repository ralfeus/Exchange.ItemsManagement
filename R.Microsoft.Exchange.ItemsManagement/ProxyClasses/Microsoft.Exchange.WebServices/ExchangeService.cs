using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class ExchangeService : BaseProxyClass
    {
        internal ExchangeService()
        {
            this.instance = ClassFactory.CreateInstance("ExchangeService", "Microsoft_Exchange_WebServices", true);
        }

        internal ExchangeService(ExchangeVersion serverVersion)
        {
            Logger.Write("Creating ExchangeService instance");
            Loader.LoadAssemblyFromResource("Microsoft_Exchange_WebService");
            var type = Loader.GetType("ExchangeService");
            var paramType = Loader.GetType("ExchangeVersion");
            this.instance = ClassFactory.CreateInstance(type.GetConstructor(new Type[] { paramType }), serverVersion);
            //this.instance = ClassFactory.CreateInstance("ExchangeService", "Microsoft_Exchange_WebServices", true, serverVersion);
            Logger.Write("Instance is created");
        }

        internal NetworkCredential Credentials
        {
            set
            {
                instance.GetType().GetProperty("Credentials").SetValue(instance, value, null);
            }
        }

        internal ImpersonatedUserId ImpersonatedUserId
        {
            set
            {
                instance.GetType().GetProperty("ImpersonateUserId").SetValue(instance, value.instance, null);
            }
        }

        internal int Timeout
        {
            get
            {
                return (int)instance.GetType().GetProperty("Timeout").GetValue(instance, null);
            }
            set
            {
                instance.GetType().GetProperty("Timeout").SetValue(instance, value, null);
            }
        }

        internal Uri Url
        {
            set
            {
                instance.GetType().GetProperty("Url").SetValue(instance, value, null);
            }
        }

        internal bool UseDefaultCredentials
        {
            set
            {
                instance.GetType().GetProperty("UseDefaultCredentials").SetValue(instance, value, null);
            }
        }
    }
}
