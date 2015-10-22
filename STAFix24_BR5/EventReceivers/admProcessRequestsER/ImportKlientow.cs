using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;

namespace admProcessRequests_EventReceiver
{
    internal class ImportKlientow
    {
        const string targetList = @"Klienci - import";

        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web, out string message)
        {
            StringBuilder sb = new StringBuilder();

            //sprawdź parametry wywołania
            SPFieldMultiChoiceValue wt = new SPFieldMultiChoiceValue(properties.ListItem["enumTypKlienta"].ToString());

            SPList list = web.Lists.TryGetList(targetList);

            for (int i = 0; i < wt.Count; i++)
            {
                switch (wt[i])
                {
                    case "KPiR":
                        Import_KPiR(web, list);
                        sb.AppendFormat(@"<li>{1} Import {0} zakończony</li>", wt[i], DateTime.Now.ToString());
                        break;
                    case "KSH":
                        Import_KSH(web, list);
                        sb.AppendFormat(@"<li>{1} Import {0} zakończony</li>", wt[i], DateTime.Now.ToString());
                        break;
                    case "Osoba fizyczna":
                        Import_OsobaFizyczna(web, list);
                        sb.AppendFormat(@"<li>{1} Import {0} zakończony</li>", wt[i], DateTime.Now.ToString());
                        break;
                    case "Firma":
                        Import_Firma(web, list);
                        sb.AppendFormat(@"<li>{1} Import {0} zakończony</li>", wt[i], DateTime.Now.ToString());
                        break;
                    default:
                        break;
                }
            }

            message = String.Format(@"<ul>{0}</ul>", sb.ToString());
        }

        private static void Import_Firma(Microsoft.SharePoint.SPWeb web, SPList list)
        {
            list.Items.Cast<SPListItem>()
                .Where(i => i["colTypRekordu"].ToString() == "Firma")
                .ToList()
                .ForEach(item =>
                {
                    string nazwaSkrocona = item["colNazwaSkrocona"] != null ? item["colNazwaSkrocona"].ToString() : string.Empty;
                    string nazwa = item["colNazwa"] != null ? item["colNazwa"].ToString() : string.Empty;
                    int klientId = tabKlienci.Get_KlientByNazwaSkrocona(web, nazwaSkrocona);

                    if (!string.IsNullOrEmpty(nazwa)
                        && !string.IsNullOrEmpty(nazwaSkrocona)
                        && klientId > 0)
                    {
                        int firmaId = tabKlienci.Get_FirmaByNazwa(web, nazwaSkrocona, nazwa);

                        if (firmaId == 0)
                        {
                            firmaId = tabKlienci.AddNew_Firma_Klient(web, item, klientId);

                            item["selKlient"] = firmaId;
                            item.SystemUpdate();
                        }

                    }
                });
        }

        private static void Import_OsobaFizyczna(Microsoft.SharePoint.SPWeb web, SPList list)
        {
            list.Items.Cast<SPListItem>()
                .Where(i => i["colTypRekordu"].ToString() == "Osoba fizyczna")
                .ToList()
                .ForEach(item =>
                {
                    string nazwaSkrocona = item["colNazwaSkrocona"] != null ? item["colNazwaSkrocona"].ToString() : string.Empty;
                    string pesel = item["colPESEL"] != null ? item["colPESEL"].ToString() : string.Empty;
                    int klientId = tabKlienci.Get_KlientByNazwaSkrocona(web, nazwaSkrocona);

                    if (!string.IsNullOrEmpty(pesel)
                        && !string.IsNullOrEmpty(nazwaSkrocona)
                        && klientId > 0)
                    {
                        int osobaId = tabKlienci.Get_OsobaFizycznaByPesel(web, nazwaSkrocona, pesel);

                        if (osobaId == 0)
                        {
                            osobaId = tabKlienci.AddNew_OsobaFizyczna_Klient(web, item, klientId);

                            item["selKlient"] = osobaId;
                            item.SystemUpdate();
                        }
                    }

                });
        }

        private static void Import_KSH(Microsoft.SharePoint.SPWeb web, SPList list)
        {
            list.Items.Cast<SPListItem>()
                .Where(i => i["colTypRekordu"].ToString() == "KSH")
                .ToList()
                .ForEach(item =>
                {
                   string nazwaSkrocona = item["colNazwaSkrocona"] != null ? item["colNazwaSkrocona"].ToString() : string.Empty;

                    if (!string.IsNullOrEmpty(nazwaSkrocona))
                    {
                        int klientId = tabKlienci.Get_KlientByNazwaSkrocona(web, nazwaSkrocona);
                        if (klientId == 0)
                        {
                            klientId = tabKlienci.AddNew_KSH_Klient(web, nazwaSkrocona, item);
                            item["selKlient"] = klientId;
                            item.SystemUpdate();
                        }
                    }
                });
        }

        private static void Import_KPiR(Microsoft.SharePoint.SPWeb web, SPList list)
        {
            list.Items.Cast<SPListItem>()
                .Where(i => i["colTypRekordu"].ToString() == "KPiR")
                .ToList()
                .ForEach(item =>
                {
                    string nazwaSkrocona = item["colNazwaSkrocona"] != null ? item["colNazwaSkrocona"].ToString() : string.Empty;

                    if (!string.IsNullOrEmpty(nazwaSkrocona))
                    {
                        int klientId = tabKlienci.Get_KlientByNazwaSkrocona(web, nazwaSkrocona);
                        if (klientId == 0)
                        {
                            klientId = tabKlienci.AddNew_KPiR_Klient(web, nazwaSkrocona, item);
                            item["selKlient"] = klientId;
                            item.SystemUpdate();
                        }
                    }
                });
        }
    }
}
