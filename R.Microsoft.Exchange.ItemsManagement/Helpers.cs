using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Win32;

namespace R.Microsoft.Exchange.ItemsManagement
{
    static class Helpers
    {
        static ExchangeVersion GetExchangeVersion()
        {
            RegistryKey msKey;
            try {
                msKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft");
            }
            catch (Exception e) {
                throw new Exception("Cannot access registry. Exchange version cannot be defined", e);
            }
            try
            {
                msKey.OpenSubKey("Exchange\\v8.0");
                return ExchangeVersion.Exchange2007_SP1;
            }
            catch { }
            try
            {
                msKey.OpenSubKey("ExchangeServer\\v14");
                return ExchangeVersion.Exchange2010;
            }
            catch
            {
                throw new Exception("No Exchange Server is installed on this computer");
            }
        }

        public static string GetExchangeBinariesDirectory()
        {
            return Helpers.GetExchangeBinariesDirectory(Helpers.GetExchangeVersion());
        }

        static string GetExchangeBinariesDirectory(ExchangeVersion exchangeVersion)
        {
            string exchangeBinariesPath = "";
            if (exchangeVersion == ExchangeVersion.Exchange2007_SP1)
            {
                var setupKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Exchange\\v8.0\\Setup");
                exchangeBinariesPath = setupKey.GetValue("MsiInstallPath").ToString() + "bin\\";
            } 
            else if ((exchangeVersion == ExchangeVersion.Exchange2010) ||
                (exchangeVersion == ExchangeVersion.Exchange2010_SP1) ||
                (exchangeVersion == ExchangeVersion.Exchange2010_SP2))
            {
                exchangeBinariesPath = Environment.GetEnvironmentVariable("ExchangeInstallPath");
                if (exchangeBinariesPath != null)
                    exchangeBinariesPath += "bin\\";
                else
                    throw new Exception("Exchange Server is not installed on this computer.");
            }
            return exchangeBinariesPath;
        }

        public static EmailAddress NormalizeEmailAddress(string address)
        {
            var parts = address.Split(new char[] { '<', '>' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2)
                return new EmailAddress(parts[0], parts[1]);
            else
                return new EmailAddress(address);
        }
    }
}
