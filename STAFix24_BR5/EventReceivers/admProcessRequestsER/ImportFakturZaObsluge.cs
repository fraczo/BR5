using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;
using BLL.Models;

namespace admProcessRequests_EventReceiver
{
    internal class ImportFakturZaObsluge
    {
        const string targetList = @"Faktury za obsługę - import"; //"intFakturyZaObsluge";

        internal static void Execute(Microsoft.SharePoint.SPItemEventProperties properties, Microsoft.SharePoint.SPWeb web)
        {
            SPListItem sItem = properties.ListItem;
            int okresId = new SPFieldLookupValue(sItem["selOkres"].ToString()).LookupId;

            SPList list = web.Lists.TryGetList(targetList);

            list.Items.Cast<SPListItem>()
                .ToList()
                .ForEach(item =>
                {
                    Import_DaneOFakturze(web, item, okresId);
                });

        }

        private static void Import_DaneOFakturze(SPWeb web, SPListItem item, int okresId)
        {
            int klientId = 0;

            string nazwaSkrocona = item["_Klient"] != null ? item["_Klient"].ToString().Trim() : string.Empty;

            if (!String.IsNullOrEmpty(nazwaSkrocona))
            {
                klientId = tabKlienci.Get_KlientId(item.Web, nazwaSkrocona);
            }

            if (klientId > 0)
            {
                //item["Title"] = tabKlienci.Get_KlientById(item.Web, klientId).Title;
                item["Title"] = String.Empty;
                item["selKlient"] = klientId;
                item["selOkres"] = okresId;

                DateTime dataWystawienia = item["cDataWystawienia"] != null ? DateTime.Parse(item["cDataWystawienia"].ToString()) : new DateTime();
                Klient iok = new Klient(web, klientId);
                DateTime terminPlatnosci = new DateTime();
                terminPlatnosci = dataWystawienia.AddDays(iok.TerminPlatnosci);
                item["colBR_TerminPlatnosci"] = terminPlatnosci;

                item.SystemUpdate();

                int zadanieId = tabZadania.Get_NumerZadaniaBR(item.Web, klientId, okresId);

                if (zadanieId > 0)
                {
                    item["selZadanie"] = zadanieId;

                    string numerFaktury = item["colBR_NumerFaktury"] != null ? item["colBR_NumerFaktury"].ToString() : string.Empty;
                    double wartoscDoZaplaty = item["colBR_WartoscDoZaplaty"] != null ? Double.Parse(item["colBR_WartoscDoZaplaty"].ToString()) : 0;



                    tabZadania.Update_InformacjeOWystawionejFakturze(web, zadanieId, numerFaktury, wartoscDoZaplaty, terminPlatnosci, dataWystawienia);

                    item.SystemUpdate();
                }
            }
            else
            {
                item["Title"] = "niezgodna nazwa pliku";
                item["selKlient"] = 0;
                item["selOkres"] = 0;
                item["selZadanie"] = 0;

                item.SystemUpdate();
            }
        }
    }
}
