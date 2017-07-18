using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement
{
    class Snapin : PSSnapIn
    {
        public override string Description
        {
            get { return "Cmdlets for Exchange items management"; }
        }

        public override string Name
        {
            get { return "R.Microsoft.Exchange.ItemsManagement"; }
        }

        public override string Vendor
        {
            get { return "Rasoft"; }
        }
    }
}
