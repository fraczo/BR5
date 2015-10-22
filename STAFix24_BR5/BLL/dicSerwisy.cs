using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class dicSerwisy
    {
        const string targetList = "Serwisy";

        internal static int Get_IdByKod(Microsoft.SharePoint.SPWeb web, string serwisName)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.Title == serwisName)
                .FirstOrDefault();
            if (item!=null)
            {
                return item.ID;
            }

            return 0;

        }
    }
}
