using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace EventReceivers.admProcessRequestsER
{
    class ObslugaZadan
    {
        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web)
        {
            ManageTasks_WorkingHours(properties);
            ManageTasks(properties);
        }

        private static void ManageTasks_WorkingHours(Microsoft.SharePoint.SPItemEventProperties properties)
        {

            TimeSpan startTS = TimeSpan.Parse(BLL.admSetup.GetValue(properties.Web, "PROC_TASK_START"));
            TimeSpan endTS = TimeSpan.Parse(BLL.admSetup.GetValue(properties.Web, "PROC_TASK_END"));
            TimeSpan currentTime = DateTime.Now.TimeOfDay;

            if (currentTime.CompareTo(startTS) < 0 || currentTime.CompareTo(endTS) > 0)
            {
                //zatrzymaj procedurę jeżeli aktualny czas nie mieści się w dozwolonym przedziale
                return;
                //Manage_ProsbaOPrzeslanieWyciaguBankowego(properties);
                //Manage_ProsbaODokumenty(properties);
                //Manage_AudytowaneZadania(properties);
            }
        }

        private static void ManageTasks(Microsoft.SharePoint.SPItemEventProperties properties)
        {
            // TODO:obsługa zadań nie podlegających restrykcjom czasowym

            Manage_ProsbaOPrzeslanieWyciaguBankowego(properties);
            Manage_ProsbaODokumenty(properties);
        }

        private static void Manage_ProsbaOPrzeslanieWyciaguBankowego(Microsoft.SharePoint.SPItemEventProperties properties)
        {
            List<SPListItem> list = BLL.tabZadania.Get_ActiveTasksByContentType(properties.Web, "Prośba o przesłanie wyciągu bankowego");
            foreach (SPListItem item in list)
            {
                Set_StatusZadania(item, "Gotowe");
            }
        }

        private static void Set_StatusZadania(SPListItem item, string status)
        {
            if (item["enumStatusZadania"]!=null)
            {
                item["enumStatusZadania"] = status;
                item.SystemUpdate();
            }
        }

        private static void Manage_ProsbaODokumenty(SPItemEventProperties properties)
        {
            List<SPListItem> list = BLL.tabZadania.Get_ActiveTasksByContentType(properties.Web, "Prośba o dokumenty");
            foreach (SPListItem item in list)
            {
                Set_StatusZadania(item, "Gotowe");
            }
        } 
    }
}
