using System;
using System.Linq;
using System.Management.Automation;
using Microsoft.Exchange.WebServices.Data;

namespace R.Microsoft.Exchange.ItemsManagement.Cmdlets
{
    [Cmdlet("New", "MailboxItem")]
    public class NewMailboxItem : BaseCmdlet
    {
        #region Parameters
        #region Common Parameters
        [Parameter]
        public string Body;

        [Parameter]
        public string Folder;

        [Parameter(ParameterSetName = "Appointment", Position = 2)]
        [Parameter(ParameterSetName = "Mail", Position = 3)]
        public string Subject;
        #endregion
        #region Appointment
        [Parameter(
            Mandatory = true,
            ParameterSetName = "Appointment",
            Position = 1)]
        public DateTime Start;
        #endregion
        #region Mail Parameters
        [Parameter(
            Mandatory = true,
            ParameterSetName = "Mail",
            Position = 2)]
        [Alias("To")]
        [ValidatePattern("<?\\S+@[\\w\\d\\-]+(\\.[\\w\\d\\-]+)*>?$")]
        public string[] Recipients;

        [Parameter(
            HelpMessageResourceId = "HelpNewMailboxItemSend",
            ParameterSetName = "Mail")]
        public SwitchParameter Send;

        [Parameter(
            Mandatory = true,
            ParameterSetName = "Mail",
            Position = 1)]
        [Alias("From")]
        [ValidatePattern("<?\\S+@[\\w\\d\\-]+(\\.[\\w\\d\\-]+)*>?$")]
        public string Sender;

        #endregion
        #endregion

        #region Overrides
        protected override void BeginProcessing()
        {
            base.BeginProcessing();
            this.ewsWrapper = new EWSWrapper(this.EWSUrl, this.Timeout, this.serverVersion);
        }

        protected override void ProcessRecord()
        {
            base.ProcessRecord();
            foreach (var upn in this.GetMailboxesUserPrincipalNames(this.Mailbox))
            {
                var destinationFolder = this.Folder != null ? this.ewsWrapper.GetFolder(upn, this.Folder) : null;
                switch (this.ParameterSetName)
                {
                    case "Appointment":
                        if (this.Folder == null)
                            destinationFolder = ewsWrapper.GetFolder(upn, WellKnownFolderName.Calendar);
                        break;
                    case "Mail":
                        if (this.Folder == null)
                            destinationFolder = ewsWrapper.GetFolder(upn, WellKnownFolderName.Inbox);
                        this.WriteObject(
                            this.ewsWrapper.CreateMailItem(
                                upn, 
                                destinationFolder, 
                                Helpers.NormalizeEmailAddress(this.Sender),
                                this.Recipients.Select((recipient) => { return Helpers.NormalizeEmailAddress(recipient); }), 
                                this.Subject, 
                                this.Body,
                                this.Send));
                        break;
                }
            }
        }
        #endregion
    }
}
