using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;
using admProcessRequests_EventReceiver;
using admProcessRequests_EventReceiver.admProcessRequestsER;
using EventReceivers.admProcessRequestsER;

namespace admProcessRequests_EventReceiver
{
    internal class GeneratorFormatekRozliczeniowych
    {
        /// <summary>
        /// Generowanie formatek rozliczeniowych dla klienta
        /// </summary>
        /// <param name="properties"></param>
        internal static void Execute_GenFormRozlK(SPItemEventProperties properties, SPWeb web)
        {
            StringBuilder msg = new StringBuilder();

            SPListItem item = properties.ListItem;

            int okresId = new SPFieldLookupValue(item["selOkres"].ToString()).LookupId;
            int klientId = new SPFieldLookupValue(item["selKlient"].ToString()).LookupId;

            if (okresId > 0 && klientId > 0)
            {
                SPListItem klient = tabKlienci.Get_KlientById(web, klientId);
                if (klient != null && klient["enumStatus"] != null && klient["enumStatus"].ToString() == "Aktywny")
                {
                    msg.AppendFormat(@"<li>klient# {0} {1}</li>",
                        klient.ID.ToString(),
                        BLL.Tools.Get_Text(klient, "_NazwaPrezentowana"));

                    switch (klient.ContentType.Name)
                    {
                        case "KPiR":
                            ZUS_Forms.Create(web, klientId, okresId);
                            PD_Forms.Create(web, klientId, okresId);
                            VAT_Forms.Create(web, klientId, okresId);
                            BR_Forms.Create(web, klientId, okresId);
                            Reminder_Forms.Create(web, klientId, okresId);
                            break;
                        case "KSH":
                            ZUS_Forms.Create(web, klientId, okresId);
                            PDS_Forms.Create(web, klientId, okresId);
                            VAT_Forms.Create(web, klientId, okresId);
                            BR_Forms.Create(web, klientId, okresId);
                            Reminder_Forms.Create(web, klientId, okresId);
                            break;
                        case "Firma":
                            PDS_Forms.Create(web, klientId, okresId);
                            break;
                        case "Osoba fizyczna":
                            ZUS_Forms.Create(web, klientId, okresId);
                            PD_Forms.Create(web, klientId, okresId);
                            PDS_Forms.Create(web, klientId, okresId);
                            VAT_Forms.Create(web, klientId, okresId);
                            break;

                        default:
                            break;
                    }
                }
            }

            // info o zakończeniu procesu
            string bodyHTML = string.Empty;

            if (msg.Length > 0)
            {
                bodyHTML = string.Format(@"<ul>{0}</ul>", msg.ToString());
            }

            string subject = string.Format(@"Generowanie formatek rozliczeniowych dla klienta");
            SPEmail.EmailGenerator.SendProcessEndConfirmationMail(subject, bodyHTML, web, item);

        }


        /// <summary>
        /// Generowanie formatek rozliczeniowych dla wszystkich aktywnych klientów
        /// </summary>
        /// <param name="properties"></param>
        internal static void Execute_GenFormRozl(SPItemEventProperties properties, SPWeb web)
        {
            StringBuilder msg = new StringBuilder();

            StringBuilder sb = new StringBuilder();

            //sprawdź parametry wywołania
            SPFieldMultiChoiceValue wt = new SPFieldMultiChoiceValue(properties.ListItem["enumTypKlienta"].ToString());
            int okresId = new SPFieldLookupValue(properties.ListItem["selOkres"].ToString()).LookupId;
            SPListItem item = properties.ListItem;

            for (int i = 0; i < wt.Count; i++)
            {
                if (okresId > 0)
                {
                    string typKlienta = wt[i];

                    Array klienci = tabKlienci.Get_AktywniKlienci_Serwis(web, typKlienta);

                    bool createKK = Get_Flag(item, "colDodajKartyKontrolne");

                    switch (typKlienta)
                    {
                        case "KPiR":
                            ZUS_Forms.Create(web, klienci, okresId, createKK);
                            PD_Forms.Create(web, klienci, okresId, createKK);
                            VAT_Forms.Create(web, klienci, okresId, createKK);
                            BR_Forms.Create(web, klienci, okresId);
                            Reminder_Forms.Create(web, klienci, okresId);
                            break;
                        case "KSH":
                            ZUS_Forms.Create(web, klienci, okresId, createKK);
                            PDS_Forms.Create(web, klienci, okresId, createKK);
                            VAT_Forms.Create(web, klienci, okresId, createKK);
                            BR_Forms.Create(web, klienci, okresId);
                            Reminder_Forms.Create(web, klienci, okresId);
                            break;
                        case "Firma":
                            PDS_Forms.Create(web, klienci, okresId, false);
                            break;
                        case "Osoba fizyczna":
                            ZUS_Forms.Create(web, klienci, okresId, false);
                            PD_Forms.Create(web, klienci, okresId, false);
                            VAT_Forms.Create(web, klienci, okresId, false);
                            break;

                        default:
                            break;
                    }

                    //informacja dla operatora
                    foreach (SPListItem klient in klienci)
                    {
                        msg.AppendFormat(@"<li>klient# {0} {1}</li>",
                            klient.ID.ToString(),
                            BLL.Tools.Get_Text(klient, "_NazwaPrezentowana"));
                    }
                }
            }

            // info o zakończeniu procesu
            string bodyHTML = string.Empty;

            if (msg.Length > 0)
            {
                bodyHTML = string.Format(@"<ul>{0}</ul>", msg.ToString());
            }

            string subject = string.Format(@"Generowanie formatek rozliczeniowych dla klientów typu {0}",
                wt.ToString());
            SPEmail.EmailGenerator.SendProcessEndConfirmationMail(subject, bodyHTML, web, item);
        }

        #region Helpers
        private static bool Get_Flag(SPListItem item, string col)
        {
            return item[col] != null ? bool.Parse(item[col].ToString()) : false;
        }
        #endregion

    }
}
