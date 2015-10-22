using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;

namespace EventReceivers.tabDochodyWspolnikow
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class tabDochodyWspolnikow : SPItemEventReceiver
    {
       /// <summary>
       /// An item was added.
       /// </summary>
       public override void ItemAdded(SPItemEventProperties properties)
       {
           base.ItemAdded(properties);
       }

       public override void ItemUpdated(SPItemEventProperties properties)
       {
           base.ItemUpdated(properties);
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
