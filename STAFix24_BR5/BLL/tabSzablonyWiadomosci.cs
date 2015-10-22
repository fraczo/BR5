using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class tabSzablonyWiadomosci
    {
        const string targetList = "Szablony wiadomości"; //"tabSzablonyWiadomosci";

        public static int GetSzablonId(Microsoft.SharePoint.SPWeb web, string nazwaSzablonu)
        {
            SPList list = web.Lists.TryGetList(targetList);
            //if (list != null)
            //{
                SPListItem item = list.Items.Cast<SPListItem>()
                    .Where(i => i.Title == nazwaSzablonu)
                    .FirstOrDefault();

                if (item != null)
                {
                    return item.ID;
                }
            //}

            return 0;
        }
    }
}
