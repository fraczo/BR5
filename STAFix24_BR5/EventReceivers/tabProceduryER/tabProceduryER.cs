using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace EventReceivers.tabProceduryER
{
    public class tabProceduryER : SPItemEventReceiver
    {
        public override void ItemAdded(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        public override void ItemUpdated(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        private void Execute(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;

            try
            {
                BLL.Logger.LogEvent(properties.WebUrl, properties.ListItem.Title + ".OnUpdate");

                SPListItem item = properties.ListItem;

                if (item.Title.StartsWith(":"))
                {
                    item["_DISPLAY"] = item.Title;
                }
                else
                {
                    item["_DISPLAY"] = string.Format("{0}::{1}",
                     item["selGrupaProcedury"] != null ? new SPFieldLookupValue(item["selGrupaProcedury"].ToString()).LookupValue : string.Empty,
                     item.Title);
                }

                item.SystemUpdate();

            }
            catch (Exception ex)
            {
                BLL.Logger.LogEvent(properties.WebUrl, ex.ToString());
                var result = ElasticEmail.EmailGenerator.ReportError(ex, properties.WebUrl.ToString());
            }
            finally
            {
                this.EventFiringEnabled = true;
            }
        }


    }
}
