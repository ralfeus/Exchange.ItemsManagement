using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using R.Microsoft.Exchange.ItemsManagement;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.Data;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            var wrapper = new EWSWrapper(new Uri("https://test-exch2010-1.test.local/EWS/Exchange.asmx"), 60000, ExchangeVersion.Exchange2010);
            //wrapper.GetMailboxItems("ad", -1, true, BodyType.HTML);
            var items = wrapper.GetMailboxItems("ad@test.local");
            //var folder = wrapper.GetFolder("dev@test.local", WellKnownFolderName.Inbox);
            var folder = wrapper.GetFolder("ad@test.local", "Sent Items");
            var folderId = new FolderId(WellKnownFolderName.Inbox, new Mailbox("dev@test.local"));
            var result = items.First().Copy(folderId);
            var service = new ExchangeService();
            //items.First().ex
        }
    }
}
