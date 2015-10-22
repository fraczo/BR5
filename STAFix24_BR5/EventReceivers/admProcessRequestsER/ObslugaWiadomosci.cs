using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace EventReceivers.admProcessRequestsER
{
    internal class ObslugaWiadomosci
    {
        const string targetList = @"Wiadomości";

        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web)
        {
            SPList list = web.Lists.TryGetList(targetList);


            list.Items.Cast<SPListItem>()
                .Where(i => (bool)i["colCzyWyslana"] != true)
                .Where(i => i["colPlanowanaDataNadania"] == null || (i["colPlanowanaDataNadania"] != null && (DateTime)i["colPlanowanaDataNadania"] <= DateTime.Now))
                .ToList()
                .ForEach(item =>
                {
                    BLL.Workflows.StartWorkflow(item, "Obsługa wiadomości");
                });

        }
    }
}
