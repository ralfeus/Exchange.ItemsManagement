using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Text;

namespace R.Microsoft.Exchange.ItemsManagement
{
    public class Logger
    {
        private BaseCmdlet callingCmdlet;
        private ProgressRecord progressRecord;
        private static Logger instance;

        private Logger() {
            this.progressRecord = new ProgressRecord(0, "Getting mailbox items", "Getting items");
        }

        public static void Init(BaseCmdlet callingCmdlet)
        {
            Logger.instance = new Logger();
            Logger.instance.callingCmdlet = callingCmdlet;
        }

        public static void Write(object entry, LogVerbosity verbosity = LogVerbosity.Normal)
        {
            if (Logger.instance == null)
                return;
                //throw new Exception("The Logger isn't initialized. Call Logger.Init() first");
            Logger.instance.WriteEntry(entry, verbosity);
        }

        private void WriteEntry(object entry, LogVerbosity verbosity)
        {
            switch (verbosity) {
                case LogVerbosity.Debug:
                    this.callingCmdlet.WriteDebug(entry.ToString());
                    break;
                case LogVerbosity.Progress:
                    this.progressRecord.StatusDescription = entry.ToString();
                    //this.progressRecord.CurrentOperation = "Stub";
                    this.callingCmdlet.WriteProgress(this.progressRecord);
                    break;
                case LogVerbosity.SubProgress:
                    var subProgress = new ProgressRecord(1, entry.ToString(), entry.ToString());
                    subProgress.ParentActivityId = this.progressRecord.ActivityId;
                    this.callingCmdlet.WriteProgress(subProgress);
                    break;
                case LogVerbosity.Verbose:
                    this.callingCmdlet.WriteVerbose(entry.ToString());
                    break;
                case LogVerbosity.Warning:
                    this.callingCmdlet.WriteWarning(entry.ToString());
                    break;
                default:
                    this.callingCmdlet.WriteObject(entry);
                    break;
            }
        }
    }

    public enum LogVerbosity {
        Debug,
        Normal,
        Progress,
        SubProgress,
        Verbose,
        Warning
    }
}
