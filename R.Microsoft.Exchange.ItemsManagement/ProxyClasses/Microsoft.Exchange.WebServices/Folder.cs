using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement.ProxyClasses.Microsoft.Exchange.WebServices
{
    class Folder : BaseProxyClass
    {
        internal Folder(object instance)
        {
            this.instance = instance;
        }

        internal static Folder Bind(ExchangeService service, WellKnownFolderName rootFolder)
        {
            Loader.LoadAssemblyFromResource("Microsoft_Exchange_WebServices");
            Type type = Loader.GetType("Folder");
            return new Folder(type.GetMethod("Bind").Invoke(null, new object[] { service.instance, rootFolder }));
        }

        internal static Folder Bind(ExchangeService service, FolderId rootFolder)
        {
            Loader.LoadAssemblyFromResource("Microsoft_Exchange_WebServices");
            Type type = Loader.GetType("Folder");
            return new Folder(type.GetMethod("Bind").Invoke(null, new object[] { service.instance, rootFolder.instance }));
        }

        internal string DisplayName
        {
            get
            {
                return (string)this.instanceType.GetProperty("DisplayName").GetValue(this.instance, null);
            }
        }

        internal FolderId Id
        {
            get
            {
                return new FolderId(this.instance.GetType().GetProperty("Id").GetValue(this.instance, null));
            }
        }

        internal FindFoldersResults FindFolders(FolderView folderView)
        {
            return new FindFoldersResults(instance.GetType().GetMethod("FindFolders").Invoke(instance, new object[] { folderView.instance }));
        }

        internal FindItemsResults<Item> FindItems(ItemView itemView)
        {
            return new FindItemsResults<Item>(instance.GetType().GetMethod("FindItems").Invoke(instance, new object[] { itemView.instance }));
        }
    }
}
