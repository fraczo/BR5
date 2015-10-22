using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Collections;

namespace BLL
{
    public class admProcessRequests
    {

        const string targetList = "admProcessRequests";

        public static void List_Cleanup(SPWeb web, int offset)
        {
            try
            {
                SPList list = web.Lists.TryGetList(targetList);
                Array results = list.Items.Cast<SPListItem>()
                    .Where(i => i["enumStatusZlecenia"].ToString() == "Zakończony")
                    .Where(i => DateTime.Parse(i["Created"].ToString()) < DateTime.Today.AddDays(-1 * offset))
                    .ToArray();

                if (results != null)
                {
                    foreach (SPListItem item in results)
                    {
                        list.Items.DeleteItemById(item.ID);
                    }
                }
            }
            catch (Exception)
            {
                
                throw;
            }
        }
    }
}
