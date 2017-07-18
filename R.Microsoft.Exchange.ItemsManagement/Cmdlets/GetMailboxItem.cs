using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;
//using R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices;

namespace R.Microsoft.Exchange.ItemsManagement
{
    [Cmdlet("Get", "MailboxItem")]
    [CmdletBinding(
        SupportsShouldProcess = true, 
        ConfirmImpact = ConfirmImpact.Medium)]
    public class GetMailboxItem : BaseCmdlet
    {
        private IEnumerable<KeyValuePair<string, object>> filter;

        #region Filter options
        [Parameter(HelpMessageResourceId = "HelpFilterEnd")]
        [Filter]
        public DateTime End;

        [Parameter(HelpMessageResourceId = "HelpFilterItemClass")]
        [Filter]
        public string ItemClass;

        /// Gets the specified WellKnownFolderNames. Enter the folder names in a comma-separated list. Wildcards are permitted. To get all, enter a value of *. 
        [Parameter(ParameterSetName = "EMail", Position = 1)]
        [Filter]
        public string MessageId;

        [Parameter(HelpMessageResourceId = "HelpFilterStart")]
        [Filter]
        public DateTime Start;

        #endregion
        
        /// Defines whether item details must be loaded. Affects performance
        [Parameter]
        public SwitchParameter Details = false;

        [Parameter]
        PSCredential EWSCredentials;

        ///The top level folder that you want to items be retrieved from. 
        ///http://msdn.microsoft.com/en-us/library/microsoft.exchange.webservices.data.wellknownfoldername(v=exchg.80).aspx 
        [Parameter]
        [Alias("TopLevelFolderName","WellKnownFolderName")] 
        public WellKnownFolderName RootFolder = WellKnownFolderName.MsgFolderRoot;
                    
        /// Maximum number of the Exchange Items to be retrieved. 
        [Parameter(HelpMessage="Items limit to be retrieved per mailbox. Defaut: unlimited")] 
        [Alias("ResultSize")] 
        [ValidateRange(1,32768)] 
        public int Limit = -1;

        /// Defines the type of body of an item. 
        /// Default: Text 
        [Parameter]
        BodyType BodyType = BodyType.Text;

        protected override void BeginProcessing()
        {
            Logger.Write("GetMailboxItem::BeginProcessing()", LogVerbosity.Verbose);
 	        base.BeginProcessing();
            this.filter = this.GetFilter();
            this.ewsWrapper = new EWSWrapper(
                this.EWSUrl, 
                this.serverVersion, 
                this.Timeout, 
                this.RootFolder,
                this.Limit,
                this.Details,
                this.BodyType,
                this.filter);
        }

        protected override void ProcessRecord()
        {
            Logger.Write("GetMailboxItem::ProcessRecord(): Begin for " + this.Mailbox, LogVerbosity.Verbose);
            try
            {
                base.ProcessRecord();
                foreach (var upn in this.GetMailboxesUserPrincipalNames(this.Mailbox))
                {
                    var items = this.ewsWrapper.GetMailboxItems(upn);
                    foreach (var item in items)
                        this.WriteObject(item);
                }
            }
            catch (Exception e)
            {
                Logger.Write(e.Message, LogVerbosity.Warning);
                Logger.Write(e.StackTrace, LogVerbosity.Warning);
            }
        }

        public IEnumerable<KeyValuePair<string, object>> GetFilter()
        {
            var res = 
                from property
                in this.GetType().GetFields()
                where property.GetCustomAttributes(typeof(FilterAttribute), false).Count() != 0
                select new KeyValuePair<string, object>(property.Name, property.GetValue(this));
            foreach (var item in res)
                Logger.Write(String.Format("{0} => {1}", item.Key, item.Value), LogVerbosity.Verbose);
            return res.Count() == 0 ? null : res;
        }

    }
}
