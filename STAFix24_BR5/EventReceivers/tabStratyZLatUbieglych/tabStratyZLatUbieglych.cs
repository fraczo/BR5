using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace EventReceivers.tabStratyZLatUbieglych
{
    public class tabStratyZLatUbieglych : SPItemEventReceiver
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
               SPListItem item = properties.ListItem;
               BLL.Tools.Ensure_LinkColumn(item, "selKlient");
           }
           catch (Exception ex)
           {
#if DEBUG
                throw ex;
#else
               BLL.Logger.LogEvent(properties.WebUrl, ex.ToString());
               var result = ElasticEmail.EmailGenerator.ReportError(ex, properties.WebUrl.ToString());
#endif

           }
           finally
           {
               this.EventFiringEnabled = true;
           }

       }

    }
}
