﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;

namespace admProcessRequests_EventReceiver
{
    internal class PD_Forms
    {

        const string ctPD = @"Rozliczenie podatku dochodowego";

        internal static void Create(SPWeb web, Array aKlienci, int okresId, bool createKK )
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
                        case @"PD-M":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_PD_M_Form(web, item.ID, okresId);
                            break;
                        case @"PD-KW":
                            if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                            Create_PD_KW_Form(web, item.ID, okresId);
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
                        case @"PD-M":
                            Create_PD_M_Form(web, item.ID, okresId);
                            break;
                        case @"PD-KW":
                            Create_PD_KW_Form(web, item.ID, okresId);
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
        private static void Create_PD_KW_Form(SPWeb web, int klientId, int okresId)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctPD, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    tabOkresy.Get_PD_KW(web, okresId, klientId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctPD_Form(web, ctPD, klientId, okresId, key, terminPlatnosci, terminPrzekazania, true);
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
        private static void Create_PD_M_Form(SPWeb web, int klientId, int okresId)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctPD, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    DateTime terminPlatnosci;
                    DateTime terminPrzekazania;

                    tabOkresy.Get_PD_M(web, okresId, klientId, out terminPlatnosci, out terminPrzekazania);

                    tabZadania.Create_ctPD_Form(web, ctPD, klientId, okresId, key, terminPlatnosci, terminPrzekazania, false);
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
