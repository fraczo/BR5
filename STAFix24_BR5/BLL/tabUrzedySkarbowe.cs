using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    class tabUrzedySkarbowe
    {
        const string targetList = "Urzędy skarbowe"; //"dicUrzedySkarbowe";

        internal static int Get_IdByName(Microsoft.SharePoint.SPWeb web, string v)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list!=null)
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

        internal static string Get_NumerRachunkuPITById(SPWeb web, int urzadId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.GetItemById(urzadId);
            if (item != null)
            {
                return item["colPIT_Konto"] != null ? item["colPIT_Konto"].ToString() : string.Empty;
            }

            return string.Empty;
        }

        internal static string Get_NazwaUrzeduById(SPWeb web, int urzadId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list != null)
            //{
            SPListItem item = list.GetItemById(urzadId);
            if (item != null)
            {
                return item.Title;
            }
            //}

            return string.Empty;
        }

        internal static string Get_NumerRachunkuCITById(SPWeb web, int urzadId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.GetItemById(urzadId);
            if (item != null)
            {
                return item["colCIT_Konto"] != null ? item["colCIT_Konto"].ToString() : string.Empty;
            }

            return string.Empty;
        }

        internal static string Get_NumerRachunkuVATById(SPWeb web, int urzadId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.GetItemById(urzadId);
            if (item != null)
            {
                return item["colVAT_Konto"] != null ? item["colVAT_Konto"].ToString() : string.Empty;
            }

            return string.Empty;
        }
    }
}
