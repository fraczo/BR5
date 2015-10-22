using System;
using Microsoft.SharePoint;
using System.Text;

namespace EventReceivers.admProcessRequestsER
{
    class ObslugaADO
    {
        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web)
        {
            StringBuilder msg = new StringBuilder();

            SPListItem item = properties.ListItem;

            // sprawdź czy wybrana procedura jest obsługiwana
            string procName = string.Empty;
            int procId = 0;
            if (item["selProcedura"] != null)
            {
                procName = new SPFieldLookupValue(item["selProcedura"].ToString()).LookupValue;
                procId = new SPFieldLookupValue(item["selProcedura"].ToString()).LookupId;
            }

            switch (procName)
            {
                case ": Rozliczenie podatku dochodowego":
                case ": Rozliczenie podatku dochodowego spółki":
                case ": Rozliczenie podatku VAT":
                case ": Rozliczenie ZUS":
                    Array tasks = BLL.tabZadania.Get_GotoweZadaniaByProceduraId(web, procId);
                    foreach (SPListItem task in tasks)
                    {
                        //Sprawdź czy klient ma ustawiony serwis AD czy ADO
                        //w przypadku AD zablokuj automatyczną akceptację

                        if (BLL.tabKlienci.Has_ServiceById(task.Web, BLL.Tools.Get_LookupId(task, "selKlient"), "ADO"))
                        {
                            //uruchom proces zatwierdzenia
                            Update_msg(msg, procName, task);
                            BLL.WorkflowHelpers.StartWorkflow(task, "Zatwierdzenie zadania");
                        }
                    }

                    break;

                case ": Prośba o dokumenty":
                case ": Prośba o przesłanie wyciągu bankowego":
                case ": Rozliczenie z biurem rachunkowym":
                    Array tasks2 = BLL.tabZadania.Get_AktywneZadaniaByProceduraId(web, procId);
                    foreach (SPListItem task in tasks2)
                    {
                        Update_msg(msg, procName, task);
                        BLL.WorkflowHelpers.StartWorkflow(task, "Zatwierdzenie zadania");
                    }
                    break;
                default:
                    break;
            }

            // info o zakończeniu procesu
            string bodyHTML = string.Empty;

            if (msg.Length > 0)
            {
                bodyHTML = string.Format(@"<ul>{0}</ul>", msg.ToString());
            }

            string subject = string.Format(@"Automatyczne zatwierdzenie zadań typu {0}", procName);
            SPEmail.EmailGenerator.SendProcessEndConfirmationMail(subject, bodyHTML, web, item);
        }

        private static void Update_msg(StringBuilder msg, string procName, SPListItem task)
        {
            msg.AppendFormat("<li>zadanie# {0} klient: {1} procedura: {2}</li>",
                task.ID.ToString(),
                BLL.Tools.Get_LookupValue(task, "selKlient"),
                procName);
        }
    }
}
