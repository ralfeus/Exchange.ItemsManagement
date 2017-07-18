using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Text;
using Microsoft.Exchange.WebServices.Data;
//using R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices;

namespace R.Microsoft.Exchange.ItemsManagement
{
    public class EWSWrapper
    {
        private ExchangeService exchangeService;
        private int timeout;
        private string userName;
        private string password;
        private WellKnownFolderName rootFolder = WellKnownFolderName.MsgFolderRoot;
        private bool details;
        private ItemView itemView = new ItemView(Int32.MaxValue);
        private int resultSize = -1;
        private BodyType bodyType;
        private SearchFilter filter;

        /// <summary>
        /// Creates an EWSWrapper object instance
        /// </summary>
        /// <param name="ewsUrl">Specifies EWS URL</param>
        /// <param name="serverVersion">Specifies Client Access Server Exchange version</param>
        /// <param name="timeout">Specifies timeout of the request</param>
        public EWSWrapper(Uri ewsUrl, int timeout, ExchangeVersion serverVersion)
        {
            exchangeService = new ExchangeService(serverVersion);
            this.exchangeService.Url = ewsUrl;
		    ServicePointManager.ServerCertificateValidationCallback += (a, b, c, d) => { return true; };
            this.exchangeService.Timeout = timeout; 
            if ( timeout < this.exchangeService.Timeout ) {
                Logger.Write("Timeout is less than builtin default value. Define greater in case the function is not responding.", LogVerbosity.Warning);
            } 
		    if ((this.userName != null) && (this.password != null)) {
			    this.exchangeService.Credentials = new NetworkCredential(this.userName, this.password);
            } else { 
                this.exchangeService.UseDefaultCredentials = true;
                Logger.Write("Credentials aren't specified. Using logged in user ones.", LogVerbosity.Verbose);
            }        
        }

        /// <summary>
        /// Creates EWSWrapper object instance and inits it for search requests
        /// </summary>
        /// <param name="ewsUrl">Specifies EWS URL</param>
        /// <param name="serverVersion">Specifies Client Access Server Exchange version</param>
        /// <param name="timeout">Specifies timeout of the request</param>
        /// <param name="rootFolder">Specifies root folder to perform search in</param>
        /// <param name="limit">Specifies maximum items to find</param>
        /// <param name="details">Specifies whether request item details like item body</param>
        /// <param name="bodyType">Specifies item body type to request</param>
        /// <param name="filter">Specifies filter elements for the search</param>
        public EWSWrapper(Uri ewsUrl, ExchangeVersion serverVersion, int timeout, WellKnownFolderName rootFolder, int limit, 
                        bool details = false, BodyType bodyType = BodyType.HTML, IEnumerable<KeyValuePair<string, object>> filter = null)
            :this(ewsUrl, timeout, serverVersion)
        {
            this.InitForSearch(rootFolder, limit, details, bodyType, filter);
        }

        public void CreateMail(string sender, string[] recipients, string subject, string body, Folder destinationFolder)
        {
        }

        private void InitForSearch(WellKnownFolderName rootFolder, int limit, bool details,
                                                BodyType bodyType, IEnumerable<KeyValuePair<string, object>> filter)
        {

            this.rootFolder = rootFolder;
            this.details = details;
            this.filter = this.GetFilter(filter);
            this.itemView = this.GetItemView();
            this.resultSize = limit;
            this.bodyType = bodyType;
        }

        private SearchFilter GetFilter(IEnumerable<KeyValuePair<string, object>> filter)
        {
            //if (filter == null)
            //    return null;
            var searchFilterCollection = new List<SearchFilter>(filter.Count());
            foreach (var filterItem in filter)
            {
                if ((filterItem.Key == "Start") && (((DateTime)filterItem.Value).ToBinary() != 0))
                    searchFilterCollection.Add(new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeCreated, filterItem.Value));
                else if ((filterItem.Key == "End") && (((DateTime)filterItem.Value).ToBinary() != 0))
                    searchFilterCollection.Add(new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeCreated, filterItem.Value));
                else if ((filterItem.Key == "ItemClass") && (filterItem.Value != null))
                    searchFilterCollection.Add(new SearchFilter.ContainsSubstring(ItemSchema.ItemClass, filterItem.Value.ToString()));
                else if ((filterItem.Key == "MessageId") && (filterItem.Value != null))
                    searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, filterItem.Value));
            }
            if (searchFilterCollection.Count == 0)
                return null;
            return new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection.ToArray());
        }

        public Folder GetFolder(string mailbox, string folderName)
        {
            var folderChain = folderName.ToLower().Split('\\');
            var mailboxFolders = this.GetSubFolders(mailbox, WellKnownFolderName.MsgFolderRoot);
            foreach (var folder in mailboxFolders)
                if (this.IsChainValid(folder, folderChain))
                    return folder;
            throw new Exception("Folder not found");
        }

        public Folder GetFolder(string mailbox, WellKnownFolderName folderName)
        {
            this.exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, mailbox);
            return Folder.Bind(this.exchangeService, folderName);
        }

        private ItemView GetItemView()
        {
            return new ItemView(Int32.MaxValue);
        }

        public IEnumerable<Item> GetMailboxItems(string mailbox) { 
            this.exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, mailbox);
            List<Item> mailboxItems = new List<Item>();
            var limit = this.resultSize;
            try {
                Logger.Write("Getting folders from " + this.rootFolder, LogVerbosity.Verbose);
                mailboxItems = this.FindItemsInFolder(Folder.Bind(this.exchangeService, this.rootFolder), ref limit, details, bodyType);
                if (limit != 0)
                {
                    var folders = this.GetSubFolders(mailbox, this.rootFolder);
                    Logger.Write(String.Format("Found {0} folders", folders.Count()), LogVerbosity.Verbose);
                    foreach (Folder folder in folders)
                    {
                        mailboxItems.AddRange(this.FindItemsInFolder(folder, ref limit, details, bodyType));
                        if (limit == 0)
                            break;
                    }
                    //Logger.Write(mailboxItems.Count);
                }
            } 
            catch (Exception e) {
                Logger.Write(e.Message, LogVerbosity.Warning);
                Logger.Write(e.StackTrace, LogVerbosity.Warning);
            } 
            return mailboxItems;
        }

        private FindFoldersResults GetSubFolders(string mailbox, WellKnownFolderName rootFolderId) {
            this.exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, mailbox);
            try
            {
                var rootFolder = Folder.Bind(this.exchangeService, rootFolderId);
                var folderView = new FolderView(Int32.MaxValue);
                folderView.Traversal = FolderTraversal.Deep;
                return rootFolder.FindFolders(folderView);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private List<Item> FindItemsInFolder(Folder folder, ref int limit, bool details, BodyType bodyType)
        {
            //var exchangeFolder = Folder.Bind(this.exchangeService, folder.Id);
            Logger.Write(String.Format("Getting items from folder {0}. Total items in folder {1}", folder.DisplayName, folder.TotalCount), LogVerbosity.Progress);
            //var itemView = new ItemView(limit < 0 ? Int32.MaxValue : limit);
            this.itemView.PageSize = limit < 0 ? Int32.MaxValue : limit;
            this.itemView.Offset = 0;
            FindItemsResults<Item> folderItems = null;
            List<Item> items = new List<Item>();
            int item = 1;
            do
            {
                this.itemView.Offset = (folderItems == null || !folderItems.NextPageOffset.HasValue) ? 0 : folderItems.NextPageOffset.Value;
                if (this.filter == null)
                {
                    Logger.Write("EWSWrapper.FindItemsInFolder(): Searching without filter", LogVerbosity.Verbose);
                    folderItems = folder.FindItems(this.itemView);
                }
                else
                {
                    Logger.Write("EWSWrapper.FindItemsInFolder(): Searching with filter", LogVerbosity.Verbose);
                    folderItems = folder.FindItems(this.filter, this.itemView);
                }
                //Write-Verbose "received $($ExchangeItems.Count)." 
                var propertySet = new PropertySet(BasePropertySet.FirstClassProperties);
                propertySet.RequestedBodyType = bodyType;
                foreach (var folderItem in folderItems)
                {
                    Logger.Write(String.Format("Getting item {0} of {1}", item++, folder.TotalCount), LogVerbosity.SubProgress);
                    if (details)
                        try
                        {
                            folderItem.Load(propertySet);
                        }
                        catch
                        {
                            //Write-Warning "$_ : $($ExchangeItem | Format-List ConversationTopic,*Date* | Out-String)".Trim() 
                        }
                    items.Add(folderItem);
                    limit--;
                }
                Logger.Write(folderItems.MoreAvailable, LogVerbosity.Verbose);
            } while (folderItems.MoreAvailable && (limit > 0));
            Logger.Write(String.Format("{0}: {1} - done", folder.DisplayName, folderItems.TotalCount), LogVerbosity.Verbose);
            return items;
        }

        private bool IsChainValid(Folder folder, IEnumerable<string> folderChain)
        {
            if (folderChain.Last<string>() == folder.DisplayName.ToLower())
                if (folderChain.Count() == 1)
                    return true;
                else
                    return this.IsChainValid(Folder.Bind(this.exchangeService, folder.ParentFolderId), folderChain.Take<string>(folderChain.Count() - 1));
            else
                return false;

        }

        internal EmailMessage CreateMailItem(string upn, Folder destinationFolder, EmailAddress sender, IEnumerable<EmailAddress> recipients, 
            string subject, string body, bool send = false)
        {
            this.exchangeService.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.PrincipalName, upn);
            var item = new EmailMessage(this.exchangeService);
            item.From = sender;
            item.ToRecipients.AddRange(recipients);
            item.Subject = subject;
            item.Body = new MessageBody(body);
            if (send)
                item.SendAndSaveCopy(destinationFolder.Id);
            else
                item.Save(destinationFolder.Id);
            return item;
        }

        internal ServiceResponseCollection<MoveCopyItemResponse> CopyItem(Item item, Folder destinationFolder)
        {
            //destinationFolder.Id.Mailbox = new Mailbox("dev@test.local");
            FolderId id = new FolderId(WellKnownFolderName.Inbox, new Mailbox("dev@test.local"));
            return this.exchangeService.CopyItems(new ItemId[] { item.Id }, id);
        }
    }
}
