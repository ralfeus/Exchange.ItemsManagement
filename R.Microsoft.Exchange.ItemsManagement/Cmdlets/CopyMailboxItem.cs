using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;
using Microsoft.Exchange.WebServices.Data;

namespace R.Microsoft.Exchange.ItemsManagement.Cmdlets
{
    [Cmdlet("Copy", "MailboxItem")]
    public class CopyMailboxItem : BaseCmdlet
    {
        [Parameter]
        [Alias("CopyTo")]
        public string DestinationMailbox;

        [Parameter(
            Mandatory = true)]
        public string DestinationFolder;

        [Parameter]
        public Item Item;

        #region Overrides
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            this.ewsWrapper = new EWSWrapper(this.EWSUrl, this.Timeout, this.serverVersion);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            foreach (var upn in this.GetMailboxesUserPrincipalNames(this.DestinationMailbox))
            {
                var destinationFolder = this.DestinationFolder != null ? this.ewsWrapper.GetFolder(upn, this.DestinationFolder) : null;
                this.WriteObject(this.ewsWrapper.CopyItem(this.Item, destinationFolder));
            }
        }
        #endregion
    }
}
