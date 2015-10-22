using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;

namespace admProcessRequests_EventReceiver
{
    internal class ImportBuforaWiadomosci
    {
        const string targetList = @"Bufor wiadomości"; //"intBuforWiadomosci";

        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web)
        {
            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
                list.Items.Cast<SPListItem>()
                    .ToList()
                    .ForEach(item =>
                    {
                        //WorkflowHelpers.StartWorkflow(item, "wfImport_WiadomoscBufora.OnDemand");
                    });

            //}
        }
    }
}
