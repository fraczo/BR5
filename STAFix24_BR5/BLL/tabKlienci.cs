using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using System.Collections;

namespace BLL
{
    public static class tabKlienci
    {
        const string listName = "Klienci"; //"tabKlienci";

        public static Array Get_AktywniKlienci(SPWeb web)
        {
            SPList targetList = web.Lists.TryGetList(listName);
            Array result = null;

            result = targetList.Items.Cast<SPListItem>()
                .Where(i => i["enumStatus"].ToString() == "Aktywny")
                .ToArray();

            return result;
        }

        public static Array Get_AktywniKlienci_Serwis(SPWeb web, string typKlienta)
        {
            SPList targetList = web.Lists.TryGetList(listName);
            Array result = null;

            result = targetList.Items.Cast<SPListItem>()
                .Where(i => i["enumStatus"].ToString() == "Aktywny")
                .Where(i => i["ContentType"].ToString() == typKlienta)
                .Where(i => new SPFieldLookupValueCollection(i["selSewisy"].ToString()).Count > 0
                || new SPFieldLookupValueCollection(i["selSerwisyWspolnicy"].ToString()).Count > 0)
                .ToArray();

            return result;
        }

        public static Array Get_AktywniKlienci_Serwis(SPWeb web)
        {
            SPList targetList = web.Lists.TryGetList(listName);
            Array result = null;

            result = targetList.Items.Cast<SPListItem>()
                .Where(i => i["enumStatus"].ToString() == "Aktywny")
                .Where(i => new SPFieldLookupValueCollection(i["selSewisy"].ToString()).Count > 0
                || new SPFieldLookupValueCollection(i["selSerwisyWspolnicy"].ToString()).Count > 0)
                .ToArray();

            return result;
        }


        public static SPListItem Get_KlientById(SPWeb web, int klientId)
        {
            SPList targetList = web.Lists.TryGetList(listName);
            SPListItem result = null;

            result = targetList.GetItemById(klientId);

            return result;
        }

        internal static bool HasServiceAssigned(SPWeb web, int klientId, string serwisKod)
        {
            bool result = false;
            SPList targetList = web.Lists.TryGetList(listName);

            //if (targetList != null)
            //{
            SPListItem item = targetList.GetItemById(klientId);
            if (item != null)
            {
                SPFieldLookupValueCollection kody = new SPFieldLookupValueCollection(item["selSewisy"].ToString());
                foreach (SPFieldLookupValue kod in kody)
                {
                    if (kod.LookupValue == serwisKod)
                    {
                        result = true;
                        break;
                    }
                }

                kody = new SPFieldLookupValueCollection(item["selSerwisyWspolnicy"].ToString());
                foreach (SPFieldLookupValue kod in kody)
                {
                    if (kod.LookupValue == serwisKod)
                    {
                        result = true;
                        break;
                    }
                }
            }
            //}

            return result;

        }

        internal static string Get_TypKlienta(SPWeb web, int klientId)
        {
            SPList targetList = web.Lists.TryGetList(listName);

            SPListItem item = targetList.GetItemById(klientId);

            return item.ContentType.Name;
            
        }

        public static int Get_KlientId(SPWeb web, string nazwaSkrocona)
        {
            int result = 0;

            SPList list = web.Lists.TryGetList(listName);
            //if (list!=null)
            //{
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["ContentType"].ToString() == "KPiR" || i["ContentType"].ToString() == "KSH")
                .Where(i => i["colNazwaSkrocona"].ToString().ToUpper() == nazwaSkrocona.ToUpper())
                .SingleOrDefault();

            if (item != null)
            {
                result = item.ID;
            }
            //}

            return result;
        }

        public static void GetNIP(SPWeb web, int klientId, out string pesel, out string nip, out string regon, out string krs)
        {
            pesel = string.Empty;
            nip = string.Empty;
            regon = string.Empty;
            krs = string.Empty;

            SPList list = web.Lists.TryGetList(listName);
            //if (list != null)
            //{
            SPListItem item = list.GetItemById(klientId);

            if (item != null)
            {
                pesel = item["colPESEL"] != null ? item["colPESEL"].ToString() : string.Empty;
                nip = item["colNIP"] != null ? item["colNIP"].ToString() : string.Empty;
                regon = item["colRegon"] != null ? item["colRegon"].ToString() : string.Empty;
                krs = item["colKRS"] != null ? item["colKRS"].ToString() : string.Empty;
            }
            //}
        }

        public static int Get_KlientId_BestFit(SPWeb web, string dluznik)
        {
            SPList list = web.Lists.TryGetList(listName);

            //szukaj w/g pełnych nazw
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i.ContentType.Name == "KPiR" || i.ContentType.Name == "KSH")
                .Where(i => i["colNazwaSkrocona"].ToString().Trim().ToUpper() == dluznik.Trim().ToUpper())
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }

            item = list.Items.Cast<SPListItem>()
                .Where(i => i.ContentType.Name == "KPiR" || i.ContentType.Name == "KSH")
                .Where(i => i["colNazwaSkrocona"].ToString().Trim().ToUpper() + "  " + i["colMiejscowosc"].ToString().Trim().ToUpper() == dluznik.Trim().ToUpper())
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }

            return 0;
        }

        public static int Get_KlientByNazwaSkrocona(SPWeb web, string nazwaSkrocona)
        {
            SPList list = web.Lists.TryGetList(listName);
            //if (list!=null)
            //{
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["ContentType"].ToString() == "KPiR" || i["ContentType"].ToString() == "KSH")
                .Where(i => i["colNazwaSkrocona"].ToString() == nazwaSkrocona)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }
            //}

            return 0;
        }


        public static int AddNew_KPiR_Klient(SPWeb web, string nazwaSkrocona, SPListItem item)
        {
            SPList list = web.Lists.TryGetList(listName);
            //if (list!=null)
            //{
            SPListItem newItem = list.AddItem();

            newItem["enumStatus"] = "Aktywny";
            newItem["ContentType"] = "KPiR";
            CopyContent(item, newItem, "colNazwaSkrocona");
            CopyContent(item, newItem, "Title");
            CopyContent(item, newItem, "colNIP");
            CopyContent(item, newItem, "colRegon");
            CopyContent(item, newItem, "colKRS");
            CopyContent(item, newItem, "colAdres");
            CopyContent(item, newItem, "colKodPocztowy");
            CopyContent(item, newItem, "colMiejscowosc");
            CopyContent(item, newItem, "enumFormaPrawna_KPiR");
            CopyContent(item, newItem, "colFormaOpodatkowaniaPD_KPiR");
            CopyContent(item, newItem, "enumRozliczeniePD");
            //urząd skarbowy
            CopyUSLookupContent(web, item, "_UrzadSkarbowy", newItem, "selUrzadSkarbowy");

            CopyContent(item, newItem, "colFormaOpodatkowaniaVAT");
            CopyContent(item, newItem, "enumRozliczenieVAT");
            //urząd skarbowy VAT
            CopyUSLookupContent(web, item, "_UrzadSkarbowyVAT", newItem, "selUrzadSkarbowyVAT");

            CopyContent(item, newItem, "colFormaOpodakowania_ZUS");
            //oddział ZUS
            CopyZUSLookupContent(web, item, "_OddzialZUS", newItem, "selOddzialZUS");

            CopyContent(item, newItem, "colDataRozpoczeciaDzialalnosci");
            CopyContent(item, newItem, "colZatrudniaPracownikow");
            CopyContent(item, newItem, "colZUS_IWA");
            CopyContent(item, newItem, "colUwagiKadrowe");
            //termin platności
            CopyTerminPlatnosciLookupContent(web, item, "_TerminPlatnosci", newItem, "selTerminPlatnosci");

            //serwisy ***
            string[] dodatkoweSerwisy = { "RBR", "ADO", "POT", "POW-Dok", "GDW" };
            CopySerwisyLookupContent(web, item, "selSewisy", newItem, "selSewisy", dodatkoweSerwisy);

            //biuro
            CopyBiuroLookupContent(web, item, "_Biuro", newItem, "selBiuro");

            //opiekun klienta
            CopyOPLookupContent(web, item, "_OpiekunKlienta", newItem, "selOpiekunKlienta");

            //dedykowany operator podatki
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Podatki", newItem, "selDedykowanyOperator_Podatki");

            // .. kadry
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Kadry", newItem, "selDedykowanyOperator_Kadry");

            // .. audyt
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Audyt", newItem, "selDedykowanyOperator_Audyt");

            CopyContent(item, newItem, "colUwagi");

            CopyContent(item, newItem, "colOsobaDoKontaktu");
            CopyContent(item, newItem, "colRola");
            CopyContent(item, newItem, "colTelefon");
            CopyEmailContent(item, newItem, "colEmail");
            CopyContent(item, newItem, "colNotatka");

            CopyContent(item, newItem, "colOsobaDoKontaktu2");
            CopyContent(item, newItem, "colRola2");
            CopyContent(item, newItem, "colTelefon2");
            CopyEmailContent(item, newItem, "colEmail2");
            CopyContent(item, newItem, "colNotatka2");

            CopyContent(item, newItem, "colOsobaDoKontaktu3");
            CopyContent(item, newItem, "colRola3");
            CopyContent(item, newItem, "colTelefon3");
            CopyEmailContent(item, newItem, "colEmail3");
            CopyContent(item, newItem, "colNotatka3");

            newItem.SystemUpdate();

            return newItem.ID;
            //}

            //return 0;
        }

        public static int AddNew_KSH_Klient(SPWeb web, string nazwaSkrocona, SPListItem item)
        {
            SPList list = web.Lists.TryGetList(listName);
            //if (list != null)
            //{
            SPListItem newItem = list.AddItem();

            newItem["enumStatus"] = "Aktywny";
            newItem["ContentType"] = "KSH";
            CopyContent(item, newItem, "colNazwaSkrocona");
            CopyContent(item, newItem, "Title");
            CopyContent(item, newItem, "colNIP");
            CopyContent(item, newItem, "colRegon");
            CopyContent(item, newItem, "colKRS");
            CopyContent(item, newItem, "colAdres");
            CopyContent(item, newItem, "colKodPocztowy");
            CopyContent(item, newItem, "colMiejscowosc");
            CopyContent(item, newItem, "enumFormaPrawna");
            CopyContent(item, newItem, "colFormaOpodatkowaniaPD_KSH");
            CopyContent(item, newItem, "enumRozliczeniePD");
            //urząd skarbowy
            CopyUSLookupContent(web, item, "_UrzadSkarbowy", newItem, "selUrzadSkarbowy");

            CopyContent(item, newItem, "colFormaOpodatkowaniaVAT");
            CopyContent(item, newItem, "enumRozliczenieVAT");
            //urząd skarbowy VAT
            CopyUSLookupContent(web, item, "_UrzadSkarbowyVAT", newItem, "selUrzadSkarbowyVAT");

            CopyContent(item, newItem, "colFormaOpodakowania_ZUS");
            //oddział ZUS
            CopyZUSLookupContent(web, item, "_OddzialZUS", newItem, "selOddzialZUS");

            CopyContent(item, newItem, "colDataRozpoczeciaDzialalnosci");
            CopyContent(item, newItem, "colZatrudniaPracownikow");
            CopyContent(item, newItem, "colZUS_IWA");
            CopyContent(item, newItem, "colUwagiKadrowe");
            //termin platności
            CopyTerminPlatnosciLookupContent(web, item, "_TerminPlatnosci", newItem, "selTerminPlatnosci");

            //biuro
            CopyBiuroLookupContent(web, item, "_Biuro", newItem, "selBiuro");

            //serwisy ***
            string[] dodatkoweSerwisy = { "AD", "POT", "POW-Dok", "POW-WBank", "GDW", "RBR" };
            CopySerwisyLookupContent(web, item, "selSewisy", newItem, "selSewisy", dodatkoweSerwisy);

            //opiekun klienta
            CopyOPLookupContent(web, item, "_OpiekunKlienta", newItem, "selOpiekunKlienta");

            //dedykowany operator podatki
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Podatki", newItem, "selDedykowanyOperator_Podatki");

            // .. kadry
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Kadry", newItem, "selDedykowanyOperator_Kadry");

            // .. audyt
            CopyOPLookupContent(web, item, "_DedykowanyOperator_Audyt", newItem, "selDedykowanyOperator_Audyt");

            CopyContent(item, newItem, "colUwagi");

            CopyContent(item, newItem, "colOsobaDoKontaktu");
            CopyContent(item, newItem, "colRola");
            CopyContent(item, newItem, "colTelefon");
            CopyEmailContent(item, newItem, "colEmail");
            CopyContent(item, newItem, "colNotatka");

            CopyContent(item, newItem, "colOsobaDoKontaktu2");
            CopyContent(item, newItem, "colRola2");
            CopyContent(item, newItem, "colTelefon2");
            CopyEmailContent(item, newItem, "colEmail2");
            CopyContent(item, newItem, "colNotatka2");

            CopyContent(item, newItem, "colOsobaDoKontaktu3");
            CopyContent(item, newItem, "colRola3");
            CopyContent(item, newItem, "colTelefon3");
            CopyEmailContent(item, newItem, "colEmail3");
            CopyContent(item, newItem, "colNotatka3");

            newItem.SystemUpdate();

            return newItem.ID;
            //}

            //return 0;
        }

        public static int AddNew_OsobaFizyczna_Klient(SPWeb web, SPListItem item, int refId)
        {
            SPList list = web.Lists.TryGetList(listName);
            //if (list != null)
            //{
            SPListItem newItem = list.AddItem();

            newItem["enumStatus"] = "Aktywny";
            newItem["ContentType"] = "Osoba fizyczna";
            //nawiązanie do firmy
            newItem["selKlient_NazwaSkrocona"] = refId;
            newItem["colNazwaSkrocona"] = string.Empty;

            CopyContent(item, newItem, "colImie");
            CopyContent(item, newItem, "colNazwisko");
            CopyContent(item, newItem, "colPESEL");
            CopyContent(item, newItem, "colNIP");
            CopyContent(item, newItem, "colAdres");
            CopyContent(item, newItem, "colKodPocztowy");
            CopyContent(item, newItem, "colMiejscowosc");
            CopyContent(item, newItem, "colFormaOpodatkowaniaPD_KPiR");
            CopyContent(item, newItem, "enumRozliczeniePD");
            CopyContent(item, newItem, "colPD_UdzialWZysku");
            //urząd skarbowy
            CopyUSLookupContent(web, item, "_UrzadSkarbowy", newItem, "selUrzadSkarbowy");

            CopyContent(item, newItem, "colFormaOpodakowania_ZUS");
            //oddział ZUS
            CopyZUSLookupContent(web, item, "_OddzialZUS", newItem, "selOddzialZUS");

            CopyContent(item, newItem, "colUwagiKadrowe");
            //termin platności

            //biuro
            CopyBiuroLookupContent(web, item, "_Biuro", newItem, "selBiuro");

            //opiekun klienta
            CopyOPLookupContent(web, item, "_OpiekunKlienta", newItem, "selOpiekunKlienta");

            CopyContent(item, newItem, "colUwagi");


            CopyContent(item, newItem, "colTelefon");
            CopyEmailContent(item, newItem, "colEmail");

            newItem.SystemUpdate();

            return newItem.ID;
            //}

            //return 0;
        }

        public static int AddNew_Firma_Klient(SPWeb web, SPListItem item, int refId)
        {
            SPList list = web.Lists.TryGetList(listName);
            //if (list != null)
            //{
            SPListItem newItem = list.AddItem();

            newItem["enumStatus"] = "Aktywny";
            newItem["ContentType"] = "Firma";
            //nawiązanie do firmy
            newItem["selKlient_NazwaSkrocona"] = refId;
            newItem["colNazwaSkrocona"] = string.Empty;

            CopyContent(item, newItem, "colNazwa");
            CopyContent(item, newItem, "colNIP");
            CopyContent(item, newItem, "colRegon");
            CopyContent(item, newItem, "colKRS");
            CopyContent(item, newItem, "colAdres");
            CopyContent(item, newItem, "colKodPocztowy");
            CopyContent(item, newItem, "colMiejscowosc");
            CopyContent(item, newItem, "enumFormaPrawna");
            CopyContent(item, newItem, "colFormaOpodatkowaniaPD_KSH");
            CopyContent(item, newItem, "enumRozliczeniePD");

            CopyContent(item, newItem, "colPD_UdzialWZysku");

            //urząd skarbowy
            CopyUSLookupContent(web, item, "_UrzadSkarbowy", newItem, "selUrzadSkarbowy");

            CopyContent(item, newItem, "colFormaOpodakowania_ZUS");
            //oddział ZUS
            CopyZUSLookupContent(web, item, "_OddzialZUS", newItem, "selOddzialZUS");

            CopyContent(item, newItem, "colUwagiKadrowe");

            newItem.SystemUpdate();

            return newItem.ID;
            //}

            //return 0;
        }



        #region Helpers

        private static void CopyOPLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1)
        {
            string v = item[c0] != null ? item[c0].ToString() : string.Empty;
            int id = BLL.dicOperatorzy.Get_IdByName(web, v);
            newItem[c1] = id;
        }

        private static void CopySerwisyLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1, string[] dodatkoweSerwisy)
        {
            SPFieldLookupValueCollection serwisy0 = new SPFieldLookupValueCollection();
            SPFieldLookupValueCollection serwisy1 = new SPFieldLookupValueCollection();

            if (item[c0] != null)
            {
                serwisy0 = item[c0] as SPFieldLookupValueCollection;

                if (newItem[c1] != null)
                {
                    serwisy1 = item[c1] as SPFieldLookupValueCollection;
                }

                foreach (SPFieldLookupValue s0Field in serwisy0)
                {
                    int s0Id = s0Field.LookupId;
                    string s0Value = s0Field.LookupValue;

                    bool found = false;
                    foreach (SPFieldLookupValue s1Field in serwisy1)
                    {
                        if (s1Field.LookupId == s0Id)
                        {
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        serwisy1.Add(s0Field);
                    }
                }

                if (dodatkoweSerwisy.Length > 0)
                {
                    foreach (string s in dodatkoweSerwisy)
                    {
                        //get serwis id
                        int sId = BLL.dicSerwisy.Get_IdByKod(web, s);

                        if (sId > 0)
                        {
                            bool found = false;
                            foreach (SPFieldLookupValue s1Field in serwisy1)
                            {
                                if (s1Field.LookupId == sId)
                                {
                                    found = true;
                                    break;
                                }
                            }

                            if (!found)
                            {
                                serwisy1.Add(new SPFieldLookupValue(sId, s));
                            }
                        }
                    }
                }

                newItem[c1] = serwisy1;
            }
        }

        private static void CopyTerminPlatnosciLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1)
        {
            string v = item[c0] != null ? item[c0].ToString() : string.Empty;
            int id = BLL.dicTerminyPlatnosci.Get_ByValue(web, int.Parse(v));
            newItem[c1] = id;
        }

        private static void CopyBiuroLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1)
        {
            string v = item[c0] != null ? item[c0].ToString() : string.Empty;
            int id = BLL.dicBiura.Get_IdByName(web, v);
            newItem[c1] = id;
        }

        private static void CopyZUSLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1)
        {
            string v = item[c0] != null ? item[c0].ToString() : string.Empty;
            int id = BLL.dicOddzialyZUS.Get_IdByName(web, v);
            newItem[c1] = id;
        }

        private static void CopyUSLookupContent(SPWeb web, SPListItem item, string c0, SPListItem newItem, string c1)
        {
            string v = item[c0] != null ? item[c0].ToString() : string.Empty;
            int id = BLL.tabUrzedySkarbowe.Get_IdByName(web, v);
            newItem[c1] = id;
        }

        private static void CopyEmailContent(SPListItem item, SPListItem newItem, string columnName)
        {
            string email = item[columnName] != null ? item[columnName].ToString() : string.Empty;
            email = email.Replace(@"mailto:", string.Empty);
            if (IsValidEmail(email.ToString()))
            {
                newItem[columnName] = email;
            }

        }

        private static void CopyContent(SPListItem item, SPListItem newItem, string columnName)
        {
            if (item[columnName] != null)
            {
                newItem[columnName] = item[columnName];
            }
        }

        private static bool IsValidEmail(string email)
        {
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return addr.Address == email;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        public static int Get_FirmaByNazwa(SPWeb web, string nazwaSkrocona, string nazwa)
        {
            SPList list = web.Lists.TryGetList(listName);

            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["ContentType"].ToString() == "Firma")
                .Where(i => new SPFieldLookupValue(i["selKlient_NazwaSkrocona"].ToString()).LookupValue == nazwaSkrocona)
                .Where(i => i["colNazwa"].ToString() == nazwa)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }


            return 0;
        }

        public static int Get_OsobaFizycznaByPesel(SPWeb web, string nazwaSkrocona, string pesel)
        {
            SPList list = web.Lists.TryGetList(listName);

            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["ContentType"].ToString() == "Osoba fizyczna")
                .Where(i => new SPFieldLookupValue(i["selKlient_NazwaSkrocona"].ToString()).LookupValue == nazwaSkrocona)
                .Where(i => i["colPESEL"].ToString() == pesel)
                .FirstOrDefault();

            if (item != null)
            {
                return item.ID;
            }

            return 0;
        }


        public static string Get_EmailById(SPWeb web, int klientId)
        {
            SPList list = web.Lists.TryGetList(listName);
            SPListItem item = list.GetItemById(klientId);
            if (item != null)
            {
                return item["colEmail"] != null ? item["colEmail"].ToString() : string.Empty;
            }

            return string.Empty;
        }

        public static string Get_NazwaFirmyById(SPWeb web, int klientId)
        {
            BLL.Models.Klient k = new Models.Klient(web, klientId);
            return k.PelnaNazwaFirmy;

        }

        public static string Get_PelnyAdresFirmyById(SPWeb web, int klientId)
        {
            BLL.Models.Klient k = new Models.Klient(web, klientId);
            return string.Format("{0}, {1} {2}", k.Adres, k.KodPocztowy, k.Miejscowosc);
        }

        public static bool Has_ServiceById(SPWeb web,int klientId, string serviceName)
        {
            bool result = false;
            SPList list = web.Lists.TryGetList(listName);
            SPListItem item = list.GetItemById(klientId);
            
            //sprawdź Serwisy
            SPFieldLookupValueCollection serwisy = item["selSewisy"] != null ? new SPFieldLookupValueCollection(item["selSewisy"].ToString()) : null;
            if (serwisy.Count>0)
            {
                foreach (SPFieldLookupValue s in serwisy)
                {
                    if (s.LookupValue == serviceName)
                    {
                        result = true;
                        break;
                    }
                }
            }

            //Sprawdź Serwisy-Wspólnicy
            SPFieldLookupValueCollection serwisy2 = item["selSerwisyWspolnicy"] != null ? new SPFieldLookupValueCollection(item["selSerwisyWspolnicy"].ToString()) : null;
            if (serwisy2.Count > 0)
            {
                foreach (SPFieldLookupValue s in serwisy2)
                {
                    if (s.LookupValue == serviceName)
                    {
                        result = true;
                        break;
                    }
                }
            }

            return result;
        }
    }
}
