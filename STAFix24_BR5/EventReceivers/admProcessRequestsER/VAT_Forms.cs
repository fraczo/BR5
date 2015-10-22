using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;
using BLL.Models;

namespace admProcessRequests_EventReceiver
{
    public class VAT_Forms
    {
        const string ctVAT = "Rozliczenie podatku VAT";

        public static void Create(SPWeb web, Array aKlienci, int okresId, bool createKK)
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
                        case @"VAT-M":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_VAT_M_Form(web, item.ID, okresId);
                            break;
                        case @"VAT-KW":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_VAT_KW_Form(web, item.ID, okresId);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        public static void Create(SPWeb web, int klientId, int okresId)
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
                        case @"VAT-M":
                            Create_VAT_M_Form(web, item.ID, okresId);
                            break;
                        case @"VAT-KW":
                            Create_VAT_KW_Form(web, item.ID, okresId);
                            break;
                        default:
                            break;
                    }
                }
            }
        }

        /// <summary>
        /// Formatka rozliczenia kwartalnego VAT
        /// </summary>
        /// <param name="web"></param>
        /// <param name="klientId"></param>
        /// <param name="okresId"></param>
        private static void Create_VAT_KW_Form(SPWeb web, int klientId, int okresId)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctVAT, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    tabOkresy.Get_VAT_KW(web, okresId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctVAT_Form(web, ctVAT, klientId, okresId, key, terminPlatnosci, terminPrzekazania, true);
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

        //formatka rozliczenia miesięcznego VAT
        private static void Create_VAT_M_Form(SPWeb web, int klientId, int okresId)
        {

            try
            {
                string key = tabZadania.Define_KEY(ctVAT, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    tabOkresy.Get_VAT_M(web, okresId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctVAT_Form(web, ctVAT, klientId, okresId, key, terminPlatnosci, terminPrzekazania, false);
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
