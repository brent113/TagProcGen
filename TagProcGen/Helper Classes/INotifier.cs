using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TagProcGen
{
    /// <summary>Log Severity Enum</summary>
    public enum LogSeverity
    {
        /// <summary>Info</summary>
        Info = 0,
        /// <summary>Warning</summary>
        Warning = 1,
        /// <summary>Critical</summary>
        Error = 2
    }

    /// <summary>Notifier Interface</summary>
    public interface INotifier
    {
        /// <summary>Log</summary>
        void Log(string Log, string Title, LogSeverity Severity);
    }
}
