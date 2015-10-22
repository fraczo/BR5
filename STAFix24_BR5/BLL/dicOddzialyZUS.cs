using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class dicOddzialyZUS
    {
        const string targetList = "Oddziały ZUS"; //"dicOddzialyZUS";

        internal static int Get_IdByName(Microsoft.SharePoint.SPWeb web, string v)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list != null)
            //{
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.Title == v)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }
            //}

            SPListItem newItem = list.AddItem();
            newItem["Title"] = v;
            newItem.SystemUpdate();

            return newItem.ID;
        }

        internal static int Ensure(SPWeb web, int oddzialId)
        {
            if (oddzialId > 0)
            {
                SPList list = web.Lists.TryGetList(targetList);
                try
                {
                    SPListItem item = list.GetItemById(oddzialId);
                    if (item != null)
                    {
                        return item.ID;
                    }
                }
                catch (Exception)
                { }
            }

            return 0;
        }
    }
}
