using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;

namespace admProcessRequests_EventReceiver
{
    internal class PDS_Forms
    {

        const string ctPDS = @"Rozliczenie podatku dochodowego spółki";

        internal static void Create(SPWeb web, Array aKlienci, int okresId, bool createKK)
        {
            foreach (SPListItem item in aKlienci)
            {
                SPFieldLookupValueCollection kody;

                switch (item.ContentType.Name)
                {
                    case "Osoba fizyczna":
                    case "Firma":
                        kody = new SPFieldLookupValueCollection(item["selSerwisyWspolnicy"].ToString());
                        break;
                    default:
                        kody = new SPFieldLookupValueCollection(item["selSewisy"].ToString());
                        break;
                }

                foreach (SPFieldLookupValue kod in kody)
                {
                    switch (kod.LookupValue)
                    {
                        case @"PDS-M":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_PDS_M_Form(web, item.ID, okresId);
                            break;
                        case @"PDS-KW":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_PDS_KW_Form(web, item.ID, okresId);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        internal static void Create(SPWeb web, int klientId, int okresId)
        {
            SPListItem item = tabKlienci.Get_KlientById(web, klientId);

            if (item != null)
            {
                SPFieldLookupValueCollection kody;

                switch (item.ContentType.Name)
                {
                    case "Osoba fizyczna":
                    case "Firma":
                        kody = new SPFieldLookupValueCollection(item["selSerwisyWspolnicy"].ToString());
                        break;
                    default:
                        kody = new SPFieldLookupValueCollection(item["selSewisy"].ToString());
                        break;
                }

                foreach (SPFieldLookupValue kod in kody)
                {
                    switch (kod.LookupValue)
                    {
                        case @"PDS-M":
                            Create_PDS_M_Form(web, item.ID, okresId);
                            break;
                        case @"PDS-KW":
                            Create_PDS_KW_Form(web, item.ID, okresId);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Formatka rozliczenia kwartalnego PD
        /// </summary>
        private static void Create_PDS_KW_Form(SPWeb web, int klientId, int okresId)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctPDS, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    //terminy płatności VAT KSH jak dla KPiR
                    tabOkresy.Get_PD_KW(web, okresId, klientId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctPDS_Form(web, ctPDS, klientId, okresId, key, terminPlatnosci, terminPrzekazania, true);
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#else
                BLL.Logger.LogEvent(web.Url, ex.ToString() + " KlientId= " + klientId.ToString());
                var result = ElasticEmail.EmailGenerator.ReportError(ex, web.Url, "KlientId=" + klientId.ToString());
#endif

            }
        }

        //formatka rozliczenia miesięcznego PDS
        private static void Create_PDS_M_Form(SPWeb web, int klientId, int okresId)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctPDS, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    //terminy płatności VAT KSH jak dla KPiR
                    tabOkresy.Get_PD_M(web, okresId, klientId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctPDS_Form(web, ctPDS, klientId, okresId, key, terminPlatnosci, terminPrzekazania, false);
                }
            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#else
                BLL.Logger.LogEvent(web.Url, ex.ToString() + " KlientId= " + klientId.ToString());
                var result = ElasticEmail.EmailGenerator.ReportError(ex, web.Url, "KlientId=" + klientId.ToString());
#endif

            }
        }
    }
}
