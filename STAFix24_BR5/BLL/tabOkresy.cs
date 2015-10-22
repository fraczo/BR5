using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class tabOkresy
    {
        const string targetList = "Okresy"; //"tabOkresy";

        public static void Get_ZUS_D(SPWeb web, int okresId, bool isChorobowa, bool isPracownicy, out double skladkaSP, out double skladkaZD, out double skladkaFP, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            skladkaSP = -1;
            skladkaZD = -1;
            skladkaFP = -1;
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {

                Get_ZUS_Terminy(isPracownicy, ref terminPlatnosci, ref terminPrzekazania, item);

                //duży ZUS - odczyt składek

                if (isChorobowa)
                {
                    skladkaSP = Double.Parse(item["colZUS_D_SPC_Skladka"].ToString());
                }
                else
                {
                    skladkaSP = Double.Parse(item["colZUS_D_SP_Skladka"].ToString());
                }

                skladkaZD = Double.Parse(item["colZUS_D_ZD_Skladka"].ToString());
                skladkaFP = Double.Parse(item["colZUS_D_FP_Skladka"].ToString());
            }
            //}
        }
        public static void Get_ZUS_M(SPWeb web, int okresId, bool isChorobowa, bool isPracownicy, out double skladkaSP, out double skladkaZD, out double skladkaFP, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            skladkaSP = -1;
            skladkaZD = -1;
            skladkaFP = -1;
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {

                Get_ZUS_Terminy(isPracownicy, ref terminPlatnosci, ref terminPrzekazania, item);

                //maly ZUS - odczyt składek

                if (isChorobowa)
                {
                    skladkaSP = Double.Parse(item["colZUS_M_SPC_Skladka"].ToString());
                }
                else
                {
                    skladkaSP = Double.Parse(item["colZUS_M_SP_Skladka"].ToString());
                }

                skladkaZD = Double.Parse(item["colZUS_M_ZD_Skladka"].ToString());
                skladkaFP = Double.Parse(item["colZUS_M_FP_Skladka"].ToString());

            }
            //}
        }
        public static void Get_ZUS_ZD(SPWeb web, int okresId, bool isChorobowa, bool isPracownicy, out double skladkaZD, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {

            skladkaZD = -1;

            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {

                Get_ZUS_Terminy(isPracownicy, ref terminPlatnosci, ref terminPrzekazania, item);

                //składka zdrowotna
                skladkaZD = Double.Parse(item["colZUS_D_ZD_Skladka"].ToString());

            }
            //}
        }

        public static void Get_ZUS_PRAC(SPWeb web, int okresId, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {

                Get_ZUS_Terminy(true, ref terminPlatnosci, ref terminPrzekazania, item);

            }
        }


        private static void Get_ZUS_Terminy(bool isPracownicy, ref DateTime terminPlatnosci, ref DateTime terminPrzekazania, SPListItem item)
        {
            if (isPracownicy)
            {
                terminPlatnosci = DateTime.Parse(item["colZUS_TerminPlatnosciSkladek_ZP"].ToString());

            }
            else
            {
                terminPlatnosci = DateTime.Parse(item["colZUS_TerminPlatnosciSkladek_Be"].ToString());
            }

            int offset = int.Parse(item["ZUS_TerminPrzekazaniaWynikow_Ofs"].ToString());
            terminPrzekazania = terminPlatnosci.AddDays(offset);
        }


        public static void Get_VAT_M(SPWeb web, int okresId, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {
                terminPlatnosci = DateTime.Parse(item["colVAT_TerminPlatnosciPodatku"].ToString());
                int ofset = int.Parse(item["colVAT_TerminPrzekazaniaWynikow_"].ToString());

                terminPrzekazania = terminPlatnosci.AddDays(ofset);
            }
            //}
        }

        public static void Get_VAT_KW(SPWeb web, int okresId, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {
                terminPlatnosci = DateTime.Parse(item["colVAT_TerminPlatnosciPodatkuKW"].ToString());
                int ofset = int.Parse(item["colVAT_TerminPrzekazaniaWynikow_"].ToString());

                terminPrzekazania = terminPlatnosci.AddDays(ofset);
            }
            //}
        }

        public static void Get_PD_M(SPWeb web, int okresId, int klientId, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            //if (list != null)
            //{
            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {
                int ofset = 0;


                string typKlienta = tabKlienci.Get_TypKlienta(web, klientId);

                switch (typKlienta)
                {
                    case "KPiR":
                    case "Osoba fizyczna":

                        terminPlatnosci = DateTime.Parse(item["colPD_TerminPlatnosciPodatku"].ToString());
                        ofset = int.Parse(item["colPD_TerminPrzekazaniaWynikow_O"].ToString());
                        break;

                    case "KSH":
                    case "Firma":

                        terminPlatnosci = DateTime.Parse(item["colCIT_TerminPlatnosciPodatku"].ToString());
                        ofset = int.Parse(item["colCIT_TerminPrzekazaniaWynikow_"].ToString());
                        break;

                    default:
                        break;
                }


                terminPrzekazania = terminPlatnosci.AddDays(ofset);


            }
            //}
        }

        public static void Get_PD_KW(SPWeb web, int okresId, int klientId, out DateTime terminPlatnosci, out DateTime terminPrzekazania)
        {
            terminPlatnosci = new DateTime();
            terminPrzekazania = new DateTime();

            SPList list = web.Lists.TryGetList(targetList);

            SPListItem item = list.GetItemById(okresId);

            if (item != null)
            {
                int ofset = 0;

                string typKlienta = tabKlienci.Get_TypKlienta(web, klientId);

                switch (typKlienta)
                {
                    case "KPiR":
                    case "Osoba fizyczna":

                        terminPlatnosci = DateTime.Parse(item["colPD_TerminPlatnosciPodatkuKW"].ToString());
                        ofset = int.Parse(item["colPD_TerminPrzekazaniaWynikow_O"].ToString());
                        break;

                    case "KSH":
                    case "Firma":

                        terminPlatnosci = DateTime.Parse(item["colCIT_TerminPlatnosciPodatkuKW"].ToString());
                        ofset = int.Parse(item["colCIT_TerminPrzekazaniaWynikow_"].ToString());
                        break;

                    default:
                        break;
                }

                terminPrzekazania = terminPlatnosci.AddDays(ofset);
            }

        }


        public static DateTime Get_TerminPlatnosciByOkresId(SPWeb web, string nazwaKolumny, int okresId)
        {
            SPListItem item = Get_ItemById(web, okresId);
            return item[nazwaKolumny] != null ? DateTime.Parse(item[nazwaKolumny].ToString()) : new DateTime();
        }

        private static SPListItem Get_ItemById(SPWeb web, int okresId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.GetItemById(okresId);
            return item;
        }

        internal static DateTime Get_TerminRealizacji(SPWeb web, int okresId, string key)
        {
            SPListItem item = Get_ItemById(web, okresId);
            DateTime startDate = Get_Date(item, "colDataZakonczenia").AddDays(1);
            DateTime targetDate = startDate;
            int offset = int.Parse(BLL.admSetup.GetValue(web, key));
            if (offset > 0)
            {
                targetDate = targetDate.AddDays(offset - 1);
            }
            while (targetDate.DayOfWeek == DayOfWeek.Saturday || targetDate.DayOfWeek == DayOfWeek.Sunday)
            {
                targetDate = targetDate.AddDays(1);
            }

            //set time of day

            TimeSpan ts = TimeSpan.Parse(BLL.admSetup.GetValue(web, "REMINDER_TIME").ToString());

            return targetDate.Add(ts);
        }

        private static DateTime Get_Date(SPListItem item, string col)
        {
            return item[col] != null ? DateTime.Parse(item[col].ToString()) : new DateTime();
        }

        public static string Get_PoprzedniMiesiacSlownieById(SPWeb web, int okresId, int offset)
        {
            SPListItem item = Get_ItemById(web, okresId);

            if (item != null)
            {
                DateTime start = BLL.Tools.Get_Date(item, "colDataRozpoczecia");
                DateTime targetDate = start.AddMonths(-1 * offset);
                switch (targetDate.Month)
                {
                    case 1: return "styczeń";
                    case 2: return "luty";
                    case 3: return "marzec";
                    case 4: return "kwiecień";
                    case 5: return "maj";
                    case 6: return "czerwiec";
                    case 7: return "lipiec";
                    case 8: return "sierpień";
                    case 9: return "wrzesień";
                    case 10: return "październik";
                    case 11: return "listopad";
                    case 12: return "grudzień";
                    default:
                        break;
                }

                return string.Empty;
            }

            return string.Empty;

        }

        internal static int Get_PoprzedniOkresIdById(SPWeb web, int okresId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.GetItemById(okresId);
            if (item!=null)
            {
                DateTime d = Get_Date(item, "colDataRozpoczecia");
                DateTime endDate = d.AddDays(-1);
                
                SPListItem foundItem = list.Items.Cast<SPListItem>()
                    .Where(i => i["colDataZakonczenia"]!=null)
                    .Where(i => (DateTime.Parse(i["colDataZakonczenia"].ToString())).Equals(endDate))
                    .FirstOrDefault();

                if (foundItem != null)
                {
                    return foundItem.ID;
                }
            }
            return 0;
        }

        internal static int Get_PoprzedniOkresKwartalnyIdById(SPWeb web, int okresId)
        {
            int mNumber = 0;
            int oId = okresId;
            do
            {
                oId = Get_PoprzedniOkresIdById(web, oId);
                try
                {
                    SPListItem item = Get_ItemById(web, oId);
                    mNumber = BLL.Tools.Get_Date(item, "colDataRozpoczecia").Month;
                }
                catch (Exception)
                {}


            } while (oId > 0 && (mNumber == 3 || mNumber == 6 || mNumber == 9 || mNumber == 12));

            return oId;
        }
    }
}
