using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class dicBiura
    {
        const string targetList = "Biura obsługi";

        internal static int Get_IdByName(Microsoft.SharePoint.SPWeb web, string v)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.Title == v)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }

            SPListItem newItem = list.AddItem();
            newItem["Title"] = v;
            newItem.SystemUpdate();

            return newItem.ID;
        }
    }
}
