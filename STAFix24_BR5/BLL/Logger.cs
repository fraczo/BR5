using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint.Workflow;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;

namespace BLL
{
    public class Logger
    {
        public static void LogEvent(string subject, string body)
        {
            SPDiagnosticsService diagSvc = SPDiagnosticsService.Local;

            diagSvc.WriteTrace(0,
                new SPDiagnosticsCategory("STAFix category", TraceSeverity.Monitorable, EventSeverity.Error),
                TraceSeverity.Monitorable,
                subject.ToString() + ":  {0}",
                new object[] { body.ToString() });
        }

    }
}
