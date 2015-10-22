using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using BLL;
using BLL.Models;

namespace admProcessRequests_EventReceiver.admProcessRequestsER
{
    public class ZUS_Forms
    {
        const string ctZUS = @"Rozliczenie ZUS";

        public static void Create(SPWeb web, Array aKlienci, int okresId, bool createKK)
        {
            foreach (SPListItem item in aKlienci)
            {
                if (createKK) BLL.tabKartyKontrolne.Create_KartaKontrolna(web, item.ID, okresId);

                Create_ZUS_Forms(web, okresId, item);
            }
        }

        public static void Create(SPWeb web, int klientId, int okresId)
        {
            SPListItem item = tabKlienci.Get_KlientById(web, klientId);

            if (item != null)
            {
                Create_ZUS_Forms(web, okresId, item);
            }
        }

        #region Helpers

        private static void Create_ZUS_Forms(SPWeb web, int okresId, SPListItem item)
        {
            bool isPracownicy = false;
            if (item["colZatrudniaPracownikow0"]!=null)
	        {
		        isPracownicy =(bool)item["colZatrudniaPracownikow0"];
	        }


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
                double skladkaSP = 0;
                double skladkaZD = 0;
                double skladkaFP = 0;
                bool isChorobowa = false;
                bool isTylkoZdrowotna = false;
                DateTime terminPlatnosci = new DateTime();
                DateTime terminPrzekazania = new DateTime();

                bool found = false;

                switch (kod.LookupValue)
                {
                    case @"ZUS-D":
                        found=true;
                        tabOkresy.Get_ZUS_D(web, okresId, isChorobowa, isPracownicy, out skladkaSP, out skladkaZD, out skladkaFP, out terminPlatnosci, out terminPrzekazania);
                        break;
                    case @"ZUS-D+C":
                        found=true;
                        isChorobowa = true;
                        tabOkresy.Get_ZUS_D(web, okresId, isChorobowa, isPracownicy, out skladkaSP, out skladkaZD, out skladkaFP, out terminPlatnosci, out terminPrzekazania);
                        break;
                    case @"ZUS-M":
                        found=true;
                        tabOkresy.Get_ZUS_M(web, okresId, isChorobowa, isPracownicy, out skladkaSP, out skladkaZD, out skladkaFP, out terminPlatnosci, out terminPrzekazania);
                        break;
                    case @"ZUS-M+C":
                        found=true;
                        isChorobowa = true;
                        tabOkresy.Get_ZUS_M(web, okresId, isChorobowa, isPracownicy, out skladkaSP, out skladkaZD, out skladkaFP, out terminPlatnosci, out terminPrzekazania);
                        break;
                    case @"ZUS-ZD":
                        found=true;
                        isTylkoZdrowotna = true;
                        tabOkresy.Get_ZUS_ZD(web, okresId, isChorobowa, isPracownicy, out skladkaZD,  out terminPlatnosci, out terminPrzekazania);
                        break;
                    case @"ZUS-PRAC":
                        found = true;
                        tabOkresy.Get_ZUS_PRAC(web, okresId, out terminPlatnosci, out terminPrzekazania);
                        break;

                    default:
                        break;
                }

                if (found)
	            {
		            Create_ZUS_Form(web, item.ID, okresId, isTylkoZdrowotna , isChorobowa, isPracownicy, skladkaSP, skladkaZD, skladkaFP, terminPlatnosci, terminPrzekazania);
                    break;
                }
            }
        }

        private static void Create_ZUS_Form(SPWeb web, int klientId, int okresId,bool isTylkoZdrowotna, bool isChorobowa, bool isPracownicy, double skladkaSP, double skladkaZD, double skladkaFP, DateTime terminPlatnosci, DateTime terminPrzekazania)
        {
            try
            {
                string key = tabZadania.Define_KEY(ctZUS, klientId, okresId);
                if (tabZadania.Check_KEY_IsAllowed(key, web, 0))
                {
                    string zus_sp_konto = admSetup.GetValue(web, "ZUS_SP_KONTO");
                    string zus_zd_konto = admSetup.GetValue(web, "ZUS_ZD_KONTO");
                    string zus_fp_konto = admSetup.GetValue(web, "ZUS_FP_KONTO");

                    Klient iok = new Klient(web, klientId);

                    tabZadania.Create_ctZUS_Form(web, ctZUS, klientId, okresId, key, isTylkoZdrowotna, isChorobowa, isPracownicy, skladkaSP, skladkaZD, skladkaFP, terminPlatnosci, terminPrzekazania, zus_sp_konto, zus_zd_konto, zus_fp_konto, iok);
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

        #endregion

    }
}
