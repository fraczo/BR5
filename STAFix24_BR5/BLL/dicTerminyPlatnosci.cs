using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class dicTerminyPlatnosci
    {
        const string targetList = "Terminy płatności"; // "dicTerminyPlatnosci";

        internal static int Get_TerminPlatnosci(Microsoft.SharePoint.SPWeb web, int terminPlatnosciId)
        {
            int result = 0;

            SPList list = web.Lists.TryGetList(targetList);

                SPListItem item = list.GetItemById(terminPlatnosciId);
                if (item!=null)
                {
                    result = item["colLiczbaDni"] != null ? int.Parse(item["colLiczbaDni"].ToString()) : 0;
                }

            return result;
        }

        internal static int Get_ByValue(SPWeb web, int liczbaDni)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => int.Parse(i["colLiczbaDni"].ToString()) == liczbaDni)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }

            SPListItem newItem = list.AddItem();
            newItem["Title"] = String.Format("{0} dni", liczbaDni.ToString());
            newItem["colLiczbaDni"] = liczbaDni;
            newItem.SystemUpdate();

            return newItem.ID;

        }
    }
}
