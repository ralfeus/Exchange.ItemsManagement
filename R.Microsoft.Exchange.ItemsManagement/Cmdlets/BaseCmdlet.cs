using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Text;
//using Microsoft.Exchange.Configuration.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace R.Microsoft.Exchange.ItemsManagement
{
    public abstract class BaseCmdlet : PSCmdlet
    {
        protected EWSWrapper ewsWrapper;
        protected ExchangeVersion serverVersion;

        #region Parameters
        [Parameter(
            HelpMessage = "Url of Exchange Web Services. Typically: https://[web]mail.yourdomain.com/ews/exchange.asmx.\nDefault: Autodetect")]
        [ValidatePattern("^https?://[^/]*/ews/exchange.asmx$")]
        public Uri EWSUrl;

        [Parameter(
            Mandatory = true,
            Position = 0,
            ValueFromPipeline = true)]
        public string Mailbox;

        /// Service Timeout in MilliSeconds 
        /// Default: 30000 
        [Parameter]
        [ValidateRange(0, 86400000)]
        protected int Timeout = 300000;
        #endregion

        public BaseCmdlet()
        {
            Logger.Init(this);

            AppDomain currentDomain = AppDomain.CurrentDomain;
            currentDomain.AssemblyResolve += new ResolveEventHandler((o, args) => {
                Logger.Write("AppDomain.CurrentDomain.AssemblyResolve.~handler(): Trying to load " + args.Name, LogVerbosity.Verbose);
                var assembly = Loader.LoadAssemblyByPath(Helpers.GetExchangeBinariesDirectory() + args.Name.Substring(0, args.Name.IndexOf(",")) + ".dll");
                if (assembly != null)
                    Logger.Write("AppDomain.CurrentDomain.AssemblyResolve.~handler(): Success ", LogVerbosity.Verbose);
                return assembly;
            });
        }

        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            Logger.Write("GetMailboxItem::BeginProcessing(): Starting EwsAutoconfig()", LogVerbosity.Verbose);
            this.casConfig();
        }

        private void casConfig()
        {
            Logger.Write("GetMailboxItem.EwsAutoconfig()", LogVerbosity.Verbose);
            if (this.EWSUrl == null)
            {
                this.EWSUrl = this.GetEwsUrl();
                Logger.Write("GetMailboxItem.EwsAutoconfig(): Found EWS URL: " + this.EWSUrl, LogVerbosity.Verbose);
            }
            this.serverVersion = this.GetEwsCasServerVersion(this.EWSUrl);
            Logger.Write("GetMailboxItem.EwsAutoconfig(): Found CAS version: " + this.serverVersion, LogVerbosity.Verbose);
        }

        private ExchangeVersion GetEwsCasServerVersion(Uri ewsUrl)
        {
            Logger.Write("GetMailboxItem.GetEwsCasServerVersion()", LogVerbosity.Verbose);
            var rootDSE = new DirectoryEntry("LDAP://RootDSE");
            var configurationNamingContext = rootDSE.Properties["configurationNamingContext"][0].ToString();
            var searcher = new DirectorySearcher(
                new DirectoryEntry("LDAP://" + configurationNamingContext),
                String.Format("(&(objectClass=msExchWebServicesVirtualDirectory)(msExchInternalHostName={0}))", ewsUrl.ToString()));
            var server = searcher.FindOne().GetDirectoryEntry().Parent.Parent.Parent;
            var serialNumber = server.Properties["serialNumber"][0].ToString();
            if (serialNumber.StartsWith("Version 8"))
                return ExchangeVersion.Exchange2007_SP1;
            else if (serialNumber.StartsWith("Version 14"))
                return ExchangeVersion.Exchange2010;
            else
                return ExchangeVersion.Exchange2013;
        }

        private Uri GetEwsUrl()
        {
            Logger.Write("GetMailboxItem.GetEwsUrl()", LogVerbosity.Verbose);
            var rootDSE = new DirectoryEntry("LDAP://RootDSE");
            var configurationNamingContext = rootDSE.Properties["configurationNamingContext"][0].ToString();
            var searcher = new DirectorySearcher(new DirectoryEntry("LDAP://" + configurationNamingContext), "objectClass=msExchWebServicesVirtualDirectory");
            var wsvd = searcher.FindOne().GetDirectoryEntry();
            Logger.Write("GetMailboxItem.EwsAutoconfig(): Got EWS virtual directory", LogVerbosity.Verbose);
            return new Uri(wsvd.Properties["msExchInternalHostName"][0].ToString());
        }

        protected IEnumerable<string> GetMailboxesUserPrincipalNames(string mailbox)
        {
            var pl = Runspace.DefaultRunspace.CreateNestedPipeline("Get-Mailbox -Identity " + mailbox + " -ResultSize unlimited", false);
            return
                from psMailbox
                in pl.Invoke()
                select ((global::Microsoft.Exchange.Data.Directory.Management.Mailbox)psMailbox.BaseObject).UserPrincipalName;
        }

        //protected void Init()
        //{

        //    //var psSnapinException = new PSSnapInException();
        //    //try
        //    //{
        //    //    Logger.Write(this.ServerVersion, LogVerbosity.Verbose);
        //    //    string snapinName;
        //    //    if (this.ServerVersion == ExchangeVersion.Exchange2007_SP1)
        //    //        snapinName = "Microsoft.Exchange.Management.PowerShell.Admin";
        //    //    else if (
        //    //        (this.ServerVersion == ExchangeVersion.Exchange2010) ||
        //    //        (this.ServerVersion == ExchangeVersion.Exchange2010_SP1) ||
        //    //        (this.ServerVersion == ExchangeVersion.Exchange2010_SP2))
        //    //        snapinName = "Microsoft.Exchange.Management.PowerShell.E2010";
        //    //    else
        //    //        snapinName = "E2013";
        //    //    Runspace.DefaultRunspace.RunspaceConfiguration.AddPSSnapIn(snapinName, out psSnapinException);
        //    //}
        //    //catch (PSArgumentException exc)
        //    //{
        //    //    if (!exc.Message.Contains("already added"))
        //    //        throw exc;
        //    //}
        //}
    }
}
