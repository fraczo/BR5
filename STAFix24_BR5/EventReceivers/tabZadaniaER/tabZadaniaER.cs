using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using BLL;
using System.Text;
using BLL.Models;
using System.Globalization;
using System.Text.RegularExpressions;

namespace EventReceivers.tabZadaniaER
{
    public class tabZadaniaER : SPItemEventReceiver
    {

        const string WYSLIJ_INFORMACJE_DO_KLIENTA = "Wyślij informację do Klienta";
        const string ZATWIERDZ = "Zatwierdź";
        const string ANULUJ = "Anuluj";

        const string emptyMarker = "---";

        //szablony do informacji o załącznikach
        const string templateH = @"<table><tr valign='top'><td><div style= 'font-family: Arial, Helvetica, sans-serif; font-size: x-small; color: #808080'><strong>w załączeniu:</strong></div></td><td><ul>{0}</ul></td></tr></table>";
        const string templateR = @"<li style= 'font-family: Arial, Helvetica, sans-serif; font-size: x-small '>{0}</li>";

        #region EventHandlers

        public override void ItemAdding(SPItemEventProperties properties)
        {
            Validate(properties);
        }

        public override void ItemUpdating(SPItemEventProperties properties)
        {
            Validate(properties);
        }

        public override void ItemAdded(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        public override void ItemUpdated(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        private void Execute(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;
            this.Execute(properties.ListItem);
            this.EventFiringEnabled = true;
        }

        #endregion

        public void Execute(SPListItem item)
        {
            try
            {
                //BLL.Logger.LogEvent(properties.WebUrl, properties.ListItem.Title + ".OnUpdate");

                SPWeb web = item.Web;

                string ct = item.ContentType.Name;
                switch (ct)
                {
                    case "Prośba o dokumenty":
                    case "Prośba o przesłanie wyciągu bankowego":
                        Update_KEY(item);
                        break;
                    case "Rozliczenie ZUS":
                    case "Rozliczenie podatku dochodowego":
                    case "Rozliczenie podatku dochodowego spółki":
                    case "Rozliczenie podatku VAT":
                    case "Rozliczenie z biurem rachunkowym":
                        Update_KEY(item);
                        Update_GBW(web, item, ct);
                        break;
                    case "Zadanie":
                        Update_Zadanie(item, web);
                        break;
                    default:
                        break;
                }

                Manage_CT(item);

                //aktualizacja tytułu rekordu
                Update_Title(item);

                //aktualizacja pola user (_)
                Update_OperatorUser(item, web);

            }
            catch (Exception ex)
            {
#if DEBUG
                throw ex;
#else
                BLL.Logger.LogEvent(item.Web.ToString(), ex.ToString());
                var result = ElasticEmail.EmailGenerator.ReportError(ex, item.Web.ToString());
#endif

            }
        }

        private void Manage_PotwierdzenieOdbioruDokumentow(SPListItem item)
        {
            if (Get_FlagValue(item, "colPotwierdzenieOdbioruDokumento"))
            {
                int klientId = Get_LookupId(item, "selKlient");
                int okresId = Get_LookupId(item, "selOkres");

                if (klientId > 0 && okresId > 0) BLL.tabZadania.Complete_PrzypomnienieOWysylceDokumentow(item, klientId, okresId);

                if (Get_Flag(item, "colPotwierdzenieOdbioruDokumento")) BLL.tabKartyKontrolne.Set_PotwierdzenieOdbioruDokumentow(item.Web, klientId, okresId);
            }
        }


        #region Updates

        /// <summary>
        /// Jeżeli operator jest przypisany to w zadaniu aktualizuje pole _KontoOperatora, które przechowuje jego login
        /// dla celów filtrowania zadań w/g bieżącego użytkownika.
        /// </summary>
        private static void Update_OperatorUser(SPListItem item, SPWeb web)
        {
            if (item["selOperator"] != null)
            {
                int operatorId = new SPFieldLookupValue(item["selOperator"].ToString()).LookupId;

                int userId = BLL.dicOperatorzy.Get_UserIdById(web, operatorId);

                if (item["_KontoOperatora"] == null)
                {
                    item["_KontoOperatora"] = userId;
                    item.SystemUpdate();
                }
                else
                {
                    if (new SPFieldUserValue(web, item["_KontoOperatora"].ToString()).LookupId != userId)
                    {
                        item["_KontoOperatora"] = userId;
                        item.SystemUpdate();
                    }
                }
            }
            else
            {
                if (item["_KontoOperatora"] != null)
                {
                    item["_KontoOperatora"] = 0;
                    item.SystemUpdate();
                }

            }
        }

        private static void Update_Title(SPListItem item)
        {
            if (String.IsNullOrEmpty(item.Title))
            {
                item["Title"] = item["selProcedura"] != null ? new SPFieldLookupValue(item["selProcedura"].ToString()).LookupValue : "#" + item.ID.ToString();
                item.SystemUpdate();
            }
        }

        private static void Update_Zadanie(SPListItem item, SPWeb web)
        {
            //przypisz procedurę na podstawie tematu
            int procId = Assign_ProceduraBasedOnTitle(item, web);

            //update termin realizacji
            Assign_TerminRealizacjiBasedOnProcedura(item, web, procId);

            //update operatora
            Assign_OperatorBasedOnProcedura(item, web, procId);
        }

        private bool Update_GBW(SPWeb web, SPListItem item, string ct)
        {
            bool result = false;
            string targetFileNameLeading = "DRUK WPŁATY__";

            UsunPodobneZalaczniki(item, targetFileNameLeading);

            if (item["colDrukWplaty"] != null && (bool)item["colDrukWplaty"])
            {
                string klient = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupValue : string.Empty;
                int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

                string nadawca = Get_Nadawca(web, klient, klientId);

                string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;

                switch (ct)
                {
                    case "Rozliczenie z biurem rachunkowym":
                        result = ObslugaGBW_RozliczenieZBiuremRachunkowym(web, item, result, targetFileNameLeading, klient, okres, nadawca);
                        break;
                    case "Rozliczenie podatku VAT":
                        result = ObslugaGBW_RozliczeniePodatkuVAT(web, item, result, targetFileNameLeading, klient, okres, klientId);
                        break;
                    case "Rozliczenie podatku dochodowego":
                    case "Rozliczenie podatku dochodowego spółki":
                        result = ObslugaGBW_RozliczeniePodatkuDochodowego(web, item, result, targetFileNameLeading, klient, okres, klientId);
                        break;
                    case "Rozliczenie ZUS":
                        result = ObslugaGBW_RozliczenieZUS(web, item, result, targetFileNameLeading, klient, okres, klientId);
                        break;
                    default:
                        break;
                }
            }

            return result;
        }

        private static void Update_KEY(SPListItem item)
        {
            string key = tabZadania.Define_KEY(item);
            tabZadania.Update_KEY(item, key);
        }

        /// <summary>
        ///jeżeli zadanie nie ma przypisanego operatora a ktoś go zaczął edytować to wtedy ta osoba zostanie przypisana
        ///do zadania o ile ma zdefiniowane konto operatora.
        /// </summary>
        /// <param name="item"></param>
        private void Update_PrzypiszOperatora(SPListItem item)
        {
            //sprawdź czy zadanie było edytowane
            DateTime datCreated = item["Created"] != null ? DateTime.Parse(item["Created"].ToString()) : new DateTime();
            DateTime datModified = item["Modified"] != null ? DateTime.Parse(item["Modified"].ToString()) : new DateTime();

            if (datCreated == datModified) return;

            //sprawdź czy przypisany operator
            int operatorId = item["selOperator"] != null ? new SPFieldLookupValue(item["selOperator"].ToString()).LookupId : 0;

            if (operatorId > 0) return;

            //sprawdź czy bieżący operator jest zdefiniowane konto operatora
            SPUser currentUser = item["Editor"] != null ? new SPFieldUserValue(item.Web, item["Editor"].ToString()).User : null;
            int targetOpId = dicOperatorzy.Get_OperatorIdByLoginName(item.Web, currentUser.LoginName);

            if (targetOpId > 0)
            {
                //przypisz operatora do zadania
                item["selOperator"] = targetOpId;
                item.SystemUpdate();
            }
        }

        #endregion

        #region Obsługa GBW

        private bool ObslugaGBW_RozliczenieZUS(SPWeb web, SPListItem item, bool result, string targetFileNameLeading, string klient, string okres, int klientId)
        {
            bool wymaganyDrukWplaty = item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false;
            string opcja = item["colZUS_Opcja"] != null ? item["colZUS_Opcja"].ToString() : string.Empty;

            if (wymaganyDrukWplaty)
            {
                double kwota;
                string konto;
                string fileName;

                if (Get_FlagValue(item, "colZatrudniaPracownikow"))
                {

                    //PIT 8AR  
                    if (Get_FlagValue(item, "colZUS_PIT-8AR_Zalaczony"))
                    {
                        kwota = item["colZUS_PIT-8AR"] != null ? Double.Parse(item["colZUS_PIT-8AR"].ToString()) : 0;
                        if (kwota > 0)
                        {
                            BLL.Models.Klient iok = new BLL.Models.Klient(web, klientId);
                            konto = iok.NumerRachunkuPD;
                            int urzadId = iok.UrzadSkarbowyId;

                            BLL.Models.UrzadSkarbowy us = new BLL.Models.UrzadSkarbowy(web, urzadId);
                            string odbiorca = us.Get_NazwaOdbiorcyPrzelewu();
                            fileName = String.Format(@"{0}PIT-8AR_{1}.pdf",
                                       targetFileNameLeading,
                                       okres);

                            string numerDeklaracji = okres.Substring(2, 2) + "M" + okres.Substring(5, 2); //TODO: obsługa kwartalnych deklaracji

                            result = Generuj_DrukPD_FromZUS(web, item, klientId, okres, kwota, konto, fileName, odbiorca, numerDeklaracji, "PIT8AR", "ZRYCZ.POD.DOCH.UM.Z", iok);
                        }
                    }
                    //PIT 4R
                    if (Get_FlagValue(item, "colZUS_PIT-4R_Zalaczony"))
                    {
                        kwota = item["colZUS_PIT-4R"] != null ? Double.Parse(item["colZUS_PIT-4R"].ToString()) : 0;
                        if (kwota > 0)
                        {
                            BLL.Models.Klient iok = new BLL.Models.Klient(web, klientId);
                            konto = iok.NumerRachunkuPD;
                            int urzadId = iok.UrzadSkarbowyId;

                            BLL.Models.UrzadSkarbowy us = new BLL.Models.UrzadSkarbowy(web, urzadId);
                            string odbiorca = us.Get_NazwaOdbiorcyPrzelewu();
                            fileName = String.Format(@"{0}PIT-4R_{1}.pdf",
                               targetFileNameLeading,
                               okres);

                            string numerDeklaracji = okres.Substring(2, 2) + "M" + okres.Substring(5, 2); //TODO: obsługa kwartalnych deklaracji

                            result = Generuj_DrukPD_FromZUS(web, item, klientId, okres, kwota, konto, fileName, odbiorca, numerDeklaracji, "PIT4R", "POD.DOCH.ZA PRAC.", iok);
                        }
                    }
                }


                switch (opcja)
                {
                    case "Tylko zdrowotna":
                        //skladka zdrowotna
                        kwota = item["colZUS_ZD_Skladka"] != null ? Double.Parse(item["colZUS_ZD_Skladka"].ToString()) : 0;
                        konto = Clean_NumerRachunku(item, "colZUS_ZD_Konto");
                        fileName = String.Format(@"{0}Składka zdrowotna_{1}.pdf",
                                           targetFileNameLeading,
                                           okres);

                        result = Generuj_DrukZUS(web, item, result, okres, klientId, kwota, konto, fileName);
                        break;
                    default:


                        //fundusz pracy
                        kwota = item["colZUS_FP_Skladka"] != null ? Double.Parse(item["colZUS_FP_Skladka"].ToString()) : 0;
                        konto = Clean_NumerRachunku(item, "colZUS_FP_Konto");
                        fileName = String.Format(@"{0}ZUS 53_{1}.pdf",
                                           targetFileNameLeading,
                                           okres);

                        result = Generuj_DrukZUS(web, item, result, okres, klientId, kwota, konto, fileName);

                        //skladka zdrowotna
                        kwota = item["colZUS_ZD_Skladka"] != null ? Double.Parse(item["colZUS_ZD_Skladka"].ToString()) : 0;
                        konto = Clean_NumerRachunku(item, "colZUS_ZD_Konto");
                        fileName = String.Format(@"{0}ZUS 52_{1}.pdf",
                                            targetFileNameLeading,
                                            okres);

                        result = Generuj_DrukZUS(web, item, result, okres, klientId, kwota, konto, fileName);

                        //skladka spoleczna
                        kwota = item["colZUS_SP_Skladka"] != null ? Double.Parse(item["colZUS_SP_Skladka"].ToString()) : 0;
                        konto = Clean_NumerRachunku(item, "colZUS_SP_Konto");
                        fileName = String.Format(@"{0}ZUS 51_{1}.pdf",
                                           targetFileNameLeading,
                                           okres);

                        result = Generuj_DrukZUS(web, item, result, okres, klientId, kwota, konto, fileName);

                        break;
                }


            }

            return result;
        }

        private bool Generuj_DrukPD_FromZUS(SPWeb web, SPListItem item, int klientId, string okres, double kwota, string konto, string fileName, string odbiorca, string numerDeklaracji, string symbolFormularza, string identyfikatorZobowiazania, BLL.Models.Klient iok)
        {
            bool result = false;

            konto = Clean_NumerRachunku(konto);

            string nadawca = iok.NazwaFirmy + " " + iok.Adres + " " + iok.KodPocztowy + " " + iok.Miejscowosc;
            nadawca = nadawca.ToUpper();
            string nip = iok.NIP;
            string typIdentyfikatora = "N";

            if (konto.Length == 26 && kwota > 0 && !string.IsNullOrEmpty(fileName))
            {
                result = GeneratorDrukow.DrukWplaty.Attach_DrukWplatyPD(web, item,
                fileName,
                odbiorca,
                konto,
                kwota,
                nadawca,
                nip,
                typIdentyfikatora,
                numerDeklaracji, //15M07
                symbolFormularza, //PIT4R
                identyfikatorZobowiazania //POD.DOCH.ZA PRAC.
                );
            }
            return result;
        }



        private static bool Generuj_DrukZUS(SPWeb web, SPListItem item, bool result, string okres, int klientId, double kwota, string konto, string fileName)
        {
            konto = Clean_NumerRachunku(konto);

            if (konto.Length == 26 && kwota > 0 && !string.IsNullOrEmpty(fileName))
            {

                string typWplaty = "S";
                string numerDeklaracji = "01";
                string okresMiesiac = okres.Substring(5, 2);
                string okresRok = okres.Substring(0, 4);
                string numerDecyzji = ".";


                BLL.Models.Klient iok = new BLL.Models.Klient(web, klientId);
                string nadawca = iok.Get_NazwaNadawcyPrzelewu().Trim().ToUpper();
                string regon = iok.Regon;
                string nip = iok.NIP;

                string identyfikatorDeklaracji = String.Format("{0} {1} {2}{3} {4}",
                    typWplaty.Substring(0, 1),
                    numerDeklaracji.Substring(0, 2),
                    okresMiesiac.Substring(0, 2),
                    okresRok.Substring(0, 4),
                    numerDecyzji = numerDecyzji.Length > 15 ? numerDecyzji.Substring(0, 15) : numerDecyzji);

                result = GeneratorDrukow.DrukWplaty.Attach_DrukWplatyZUS(web, item,
                fileName,
                konto,
                kwota,
                nadawca, nip, "R", regon, identyfikatorDeklaracji);
            }
            return result;
        }

        private bool ObslugaGBW_RozliczeniePodatkuDochodowego(SPWeb web, SPListItem item, bool result, string targetFileNameLeading, string klient, string okres, int klientId)
        {
            bool wymaganyDrukWplaty = item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false;
            string ocenaWyniku = item["colPD_OcenaWyniku"] != null ? item["colPD_OcenaWyniku"].ToString() : string.Empty;
            double kwota = item["colPD_WartoscDoZaplaty"] != null ? Double.Parse(item["colPD_WartoscDoZaplaty"].ToString()) : 0;
            string konto = Clean_NumerRachunku(item, "colPD_Konto");

            if (wymaganyDrukWplaty && konto.Length == 26 && kwota > 0 && ocenaWyniku == "Dochód")
            {

                string fileName = String.Format(@"{0}Podatek dochodowy_{1}.pdf",
                    targetFileNameLeading,
                    okres);

                int urzadId = item["selUrzadSkarbowy"] != null ? new SPFieldLookupValue(item["selUrzadSkarbowy"].ToString()).LookupId : 0;
                BLL.Models.UrzadSkarbowy us = new BLL.Models.UrzadSkarbowy(web, urzadId);
                string odbiorca = us.Get_NazwaOdbiorcyPrzelewu();

                //jeżeli rozliczenie kwartalne

                string numerDeklaracji = string.Empty;

                string rozliczenie = item["enumRozliczeniePD"] != null ? item["enumRozliczeniePD"].ToString() : string.Empty;
                if (rozliczenie == "Kwartalnie")
                {
                    numerDeklaracji = okres.Substring(2, 2) + "K";

                    string m = okres.Substring(5, 2); //oznaczenie miesiąca
                    switch (m)
                    {
                        case "01":
                        case "02":
                        case "03":
                            numerDeklaracji = numerDeklaracji + "01";
                            break;
                        case "04":
                        case "05":
                        case "06":
                            numerDeklaracji = numerDeklaracji + "02";
                            break;
                        case "07":
                        case "08":
                        case "09":
                            numerDeklaracji = numerDeklaracji + "03";
                            break;
                        case "10":
                        case "11":
                        case "12":
                            numerDeklaracji = numerDeklaracji + "04";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    numerDeklaracji = okres.Substring(2, 2) + "M" + okres.Substring(5, 2);
                }

                BLL.Models.Klient iok = new BLL.Models.Klient(web, klientId);


                string nadawca = iok.Get_NazwaNadawcyPrzelewu();

                //nip z kartoteki
                string nip = iok.NIP;
                string typIdentyfikatora = "N";

                string symbolFormularza = Get_SymbolFormularzaPD(item);
                string opis = "ZAL.POD.DOCH.";


                result = GeneratorDrukow.DrukWplaty.Attach_DrukWplatyPD(web, item,
                fileName,
                odbiorca,
                konto,
                kwota,
                nadawca,
                nip, typIdentyfikatora, numerDeklaracji, symbolFormularza, opis);
            }
            return result;
        }

        private string Get_SymbolFormularzaPD(SPListItem item)
        {
            string result = string.Empty;
            string formaOpodatkowaniaPD = item["colFormaOpodatkowaniaPD"] != null ? item["colFormaOpodatkowaniaPD"].ToString() : string.Empty;
            switch (formaOpodatkowaniaPD)
            {
                case "CIT":
                    result = "CIT-8";
                    break;
                case "Zasady ogólne":
                    result = "PIT-5";
                    break;
                case "Podatek liniowy":
                    result = "PPL";
                    break;
                case "Karta podatkowa":
                    result = "KP";
                    break;
                case "Ryczałt":
                    result = "PPE";
                    break;
                default:
                    break;
            }

            return result;
        }


        /// <summary>
        /// jeżeli numer rachunku ma więcej niż 26 znaków usówa wszystkie znaki nie będące cyfrą
        /// </summary>
        private string Clean_NumerRachunku(SPListItem item, string colName)
        {
            string numerRachunku = item[colName] != null ? item[colName].ToString() : string.Empty;

            numerRachunku = Clean_NumerRachunku(numerRachunku);

            return numerRachunku;
        }

        private static string Clean_NumerRachunku(string numerRachunku)
        {
            if (numerRachunku.Length > 26)
            {
                Regex rgx = new Regex("[^0-9]");
                numerRachunku = rgx.Replace(numerRachunku, "");
            }
            return numerRachunku;
        }

        private static string Get_Nadawca(SPWeb web, string klient, int klientId)
        {
            string pesel = string.Empty;
            string nip = string.Empty;
            string regon = string.Empty;
            string krs = string.Empty;

            tabKlienci.GetNIP(web, klientId, out pesel, out nip, out regon, out krs);

            string nadawca = String.Format("{0}{1}{2}{3}{4}",
                                klient,
                                !string.IsNullOrEmpty(pesel) ? " PESEL:" + pesel : string.Empty,
                                !string.IsNullOrEmpty(nip) ? " NIP:" + nip : string.Empty,
                                !string.IsNullOrEmpty(regon) ? " REGON:" + regon : string.Empty,
                                !string.IsNullOrEmpty(krs) ? " KRS" + krs : string.Empty);
            return nadawca;
        }

        private bool ObslugaGBW_RozliczeniePodatkuVAT(SPWeb web, SPListItem item, bool result, string targetFileNameLeading, string klient, string okres, int klientId)
        {
            bool wymaganyDrukWplaty = item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false;
            string decyzja = item["colVAT_Decyzja"] != null ? item["colVAT_Decyzja"].ToString() : string.Empty;
            double kwota = item["colVAT_WartoscDoZaplaty"] != null ? Double.Parse(item["colVAT_WartoscDoZaplaty"].ToString()) : 0;
            string konto = Clean_NumerRachunku(item, "colVAT_Konto");

            if (wymaganyDrukWplaty && konto.Length == 26 && kwota > 0 && decyzja == "Do zapłaty")
            {
                string fileName = String.Format(@"{0}Podatek VAT_{1}.pdf",
                    targetFileNameLeading,
                    okres);

                int urzadId = item["selUrzadSkarbowy"] != null ? new SPFieldLookupValue(item["selUrzadSkarbowy"].ToString()).LookupId : 0;
                BLL.Models.UrzadSkarbowy us = new BLL.Models.UrzadSkarbowy(web, urzadId);
                string odbiorca = us.Get_NazwaOdbiorcyPrzelewu();

                string numerDeklaracji = string.Empty;

                string rozliczenie = item["enumRozliczenieVAT"] != null ? item["enumRozliczenieVAT"].ToString() : string.Empty;
                if (rozliczenie == "Kwartalnie")
                {
                    numerDeklaracji = okres.Substring(2, 2) + "K";

                    string m = okres.Substring(5, 2); //oznaczenie miesiąca
                    switch (m)
                    {
                        case "01":
                        case "02":
                        case "03":
                            numerDeklaracji = numerDeklaracji + "01";
                            break;
                        case "04":
                        case "05":
                        case "06":
                            numerDeklaracji = numerDeklaracji + "02";
                            break;
                        case "07":
                        case "08":
                        case "09":
                            numerDeklaracji = numerDeklaracji + "03";
                            break;
                        case "10":
                        case "11":
                        case "12":
                            numerDeklaracji = numerDeklaracji + "04";
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    numerDeklaracji = okres.Substring(2, 2) + "M" + okres.Substring(5, 2);
                }

                BLL.Models.Klient iok = new BLL.Models.Klient(web, klientId);
                string nadawca = iok.Get_NazwaNadawcyPrzelewu();

                //nip z kartoteki
                string nip = iok.NIP;
                string typIdentyfikatora = "N";

                string symbolFormularza = iok.FormaOpodatkowaniaVAT.Replace("-", "").Trim();
                string opis = "Podatek VAT";


                result = GeneratorDrukow.DrukWplaty.Attach_DrukWplatyPD(web, item,
                fileName,
                odbiorca,
                konto,
                kwota,
                nadawca, nip, typIdentyfikatora, numerDeklaracji, symbolFormularza, opis);
            }
            return result;
        }

        private bool ObslugaGBW_RozliczenieZBiuremRachunkowym(SPWeb web, SPListItem item, bool result, string targetFileNameLeading, string klient, string okres, string nadawca)
        {
            double kwota = item["colBR_WartoscDoZaplaty"] != null ? Double.Parse(item["colBR_WartoscDoZaplaty"].ToString()) : 0;
            string konto = Clean_NumerRachunku(item, "colBR_Konto");

            if (konto.Length == 26 && kwota > 0)
            {
                string fileName = String.Format(@"{0}Faktura za obsługę księgową_{1}.pdf",
                    targetFileNameLeading,
                    okres);
                //string odbiorca = admSetup.GetValue(web, "BR_NAZWA");
                string odbiorca = admSetup.Get_NazwaBiura(web);
                string numerFaktury = item["colBR_NumerFaktury"] != null ? item["colBR_NumerFaktury"].ToString() : string.Empty;
                string tytulem = String.Format("Zapłata za FV {0}", numerFaktury);

                result = GeneratorDrukow.DrukWplaty.Attach_DrukWplaty(web, item,
                fileName,
                odbiorca,
                konto,
                kwota,
                nadawca,
                tytulem);
            }
            return result;
        }

        #endregion

        private void Manage_CT(SPListItem item)
        {
            string status = item["enumStatusZadania"] != null ? item["enumStatusZadania"].ToString() : string.Empty;

            if (status != StatusZadania.Zakończone.ToString()
                && status != StatusZadania.Anulowane.ToString())
            {
                //Obsługa poleceń cmdFormatka - może zmieniać statusy zadania.
                Manage_CMD(item);

                if (item["enumStatusZadania"].ToString() == StatusZadania.Nowe.ToString())
                {
                    Update_PrzypiszOperatora(item);
                    Update_StatusZadania(item, StatusZadania.Obsługa);
                }

                string ct = item.ContentType.Name;
                switch (ct)
                {
                    case "Zadanie":
                        break;
                    case "Prośba o dokumenty":
                        break;
                    case "Prośba o przesłanie wyciągu bankowego":
                        break;
                    case "Rozliczenie z biurem rachunkowym":
                        break;
                    case "Rozliczenie podatku dochodowego":
                    case "Rozliczenie podatku dochodowego spółki":
                        Manage_PotwierdzenieOdbioruDokumentow(item);
                        break;
                    case "Rozliczenie podatku VAT":
                        break;
                    case "Rozliczenie ZUS":
                        break;

                    case "Wiadomość indywidualna":
                    case "Wiadomość indywidualna z szablonu":
                    case "Wiadomość do grupy":
                    case "Wiadomość do grupy z szablonu":
                        BLL.tabZadania.CreateMail(item);
                        break;

                    default:
                        break;
                }
            }
        }


        private void Manage_CMD(SPListItem item)
        {
            //sprawdź czy wybrana jakaś komenda
            string cmd = GetCommand(item);

            if (string.IsNullOrEmpty(cmd)) return;

            //obsługa komend
            switch (cmd)
            {
                case ZATWIERDZ:
                    Manage_CMD_Zatwierdz(item);
                    //nie usówaj informacji z pola informacje dla klienta
                    ResetCommand(item, false);
                    break;
                case WYSLIJ_INFORMACJE_DO_KLIENTA:
                    Manage_CMD_WyslijInfo(item);
                    //wyczyść informacje dla klienta po wysyłce
                    ResetCommand(item, true);
                    break;
                case ANULUJ:
                    Manage_CMD_Anuluj(item);
                    //wyczyść informacje dla klienta po wysyłce
                    ResetCommand(item, false);
                    break;
                default:
                    break;

            }
        }

        #region Manage CT

        private void Manage_CMD_Anuluj(SPListItem item)
        {
            string cmd = GetCommand(item);
            if (cmd == ANULUJ)
            {
                Update_StatusZadania(item, StatusZadania.Anulowane);
            }
        }

        private void Manage_CMD_WyslijInfo(SPListItem item)
        {
            string ct = item.ContentType.Name;

            switch (ct)
            {
                case "Zadanie":
                    Manage_CMD_WyslijInfo_Zadanie(item);
                    break;
                case "Prośba o dokumenty":
                case "Prośba o przesłanie wyciągu bankowego":
                case "Rozliczenie z biurem rachunkowym":
                case "Rozliczenie podatku dochodowego":
                case "Rozliczenie podatku VAT":
                case "Rozliczenie ZUS":
                    Manage_CMD_WyslijInfo_NoAtt(item);
                    break;
                default:
                    break;

            }
        }

        private void Manage_CMD_WyslijInfo_Zadanie(SPListItem item)
        {

            string cmd = GetCommand(item);
            string notatka = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == WYSLIJ_INFORMACJE_DO_KLIENTA
                && !string.IsNullOrEmpty(notatka))
            {
                string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;
                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, klientId);

                string kopiaDla = Get_KopiaDlaOperatora(item);

                bool KopiaDoNadawcy = true;
                bool KopiaDoBiura = false;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;
                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "EMAIL_DEFAULT_BODY.Include", out temat, out trescHTML, nadawca);
                if (item["selProcedura"] != null)
                {
                    temat = string.Format("{0} :{1}",
                        new SPFieldLookupValue(item["selProcedura"].ToString()).LookupValue,
                        item.Title);
                }
                else
                {
                    temat = item.Title;
                }
                if (!temat.StartsWith(":"))
                {
                    temat = ": " + temat.Trim();
                }

                temat = AddSygnatura(temat, item);
                temat = BLL.Tools.AddCompanyName(temat, item);

                StringBuilder sb = new StringBuilder(trescHTML);
                sb.Replace("___BODY___", notatka);
                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
            }

            ResetCommand(item, true);
        }

        /// <summary>
        /// jeżeli Editor nie jest aktualnym właścicielem zadania dodaj kopię do właściciela
        /// </summary>
        private string Get_KopiaDlaOperatora(SPListItem item)
        {
            int operatorId = Get_LookupId(item, "selOperator");
            string editorLoginName = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.LoginName;
            int editorId = BLL.dicOperatorzy.Get_OperatorIdByLoginName(item.Web, editorLoginName);
            if (operatorId > 0 && editorId != operatorId)
            {
                return BLL.dicOperatorzy.Get_EmailById(item.Web, operatorId);
            }

            return string.Empty;
        }



        private void Manage_CMD_WyslijInfo_NoAtt(SPListItem item)
        {
            string cmd = GetCommand(item);

            string notatka = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == WYSLIJ_INFORMACJE_DO_KLIENTA)
            {

                string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;
                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);

                string kopiaDla = Get_KopiaDlaOperatora(item);

                bool KopiaDoNadawcy = true; //wyślij kopię do nadawcy
                bool KopiaDoBiura = false;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;
                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "EMAIL_DEFAULT_BODY.Include", out temat, out trescHTML, nadawca);
                //nadpisz temat wiadomości
                if (item["selProcedura"] != null)
                {
                    temat = new SPFieldLookupValue(item["selProcedura"].ToString()).LookupValue;
                }
                else
                {
                    temat = item.Title;
                }
                if (!temat.StartsWith(":"))
                {
                    temat = ":" + temat.Trim();
                }

                temat = string.Format("{0} - informacja uzupełniająca", temat, item.ID.ToString());

                temat = AddSygnatura(temat, item);
                temat = BLL.Tools.AddCompanyName(temat, item);

                StringBuilder sb = new StringBuilder(trescHTML);
                sb.Replace("___BODY___", notatka);
                sb.Replace("___FOOTER___", string.Empty);
                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = new DateTime(); //wyślij natychmiast

                //nie kopiuj załączników
                BLL.tabWiadomosci.AddNew(item.Web, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
            }
        }

        private void Manage_CMD_Zatwierdz(SPListItem item)
        {
            string cmd = GetCommand(item);
            if (cmd == ZATWIERDZ)
            {
                string ct = item.ContentType.Name;

                switch (ct)
                {
                    case "Zadanie":
                        Update_StatusZadania(item, StatusZadania.Zakończone);
                        break;
                    case "Prośba o przesłanie wyciągu bankowego":
                        Manage_CMD_WyslijWynik_ProsbaOWyciagBankowy(item);
                        Update_StatusZadania(item, StatusZadania.Wysyłka);
                        break;
                    case "Prośba o dokumenty":
                        Manage_CMD_WyslijWynik_ProsbaODokumenty(item);
                        Update_StatusZadania(item, StatusZadania.Wysyłka);
                        break;
                    case "Rozliczenie ZUS":
                        if (isValidated_ZUS(item))
                        {
                            if (!isAuditRequest(item) || Get_StatusZadania(item) == StatusZadania.Gotowe.ToString()) //zatwiedzenie gotowego zadania powoduje jego zwolnienie
                            {
                                Manage_CMD_WyslijWynik_ZUS(item);
                                Update_KartaKlienta_ZUS(item);
                                Update_StatusZadania(item, StatusZadania.Wysyłka);
                            }
                            else
                            {
                                Update_StatusZadania(item, StatusZadania.Gotowe);
                            }
                        }
                        break;
                    case "Rozliczenie podatku dochodowego":
                        if (isValidated_PD(item))
                        {
                            if (!isAuditRequest(item) || Get_StatusZadania(item) == StatusZadania.Gotowe.ToString()) //zatwiedzenie gotowego zadania powoduje jego zwolnienie
                            {
                                Manage_CMD_WyslijWynik_PD(item);
                                Update_KartaKlienta_PD(item);
                                Update_StatusZadania(item, StatusZadania.Wysyłka);
                            }
                            else
                            {
                                Update_StatusZadania(item, StatusZadania.Gotowe);
                            }
                        }
                        break;
                    case "Rozliczenie podatku dochodowego spółki":
                        if (isValidated_PDS(item))
                        {
                            if (!isAuditRequest(item) || Get_StatusZadania(item) == StatusZadania.Gotowe.ToString()) //zatwiedzenie gotowego zadania powoduje jego zwolnienie
                            {
                                Manage_CMD_WyslijWynik_PDS(item);
                                Update_KartaKlienta_PDS(item);
                                Update_StatusZadania(item, StatusZadania.Wysyłka);
                            }
                            else
                            {
                                Update_StatusZadania(item, StatusZadania.Gotowe);
                            }
                        }
                        break;
                    case "Rozliczenie podatku VAT":
                        if (isValidated_VAT(item))
                        {
                            if (!isAuditRequest(item) || Get_StatusZadania(item) == StatusZadania.Gotowe.ToString()) //zatwiedzenie gotowego zadania powoduje jego zwolnienie
                            {
                                Manage_CMD_WyslijWynik_VAT(item);
                                Update_KartaKlienta_VAT(item);
                                Update_StatusZadania(item, StatusZadania.Wysyłka);
                            }
                            else
                            {
                                Update_StatusZadania(item, StatusZadania.Gotowe);
                            }
                        }
                        break;

                    case "Rozliczenie z biurem rachunkowym":
                        if (isValidated_RBR(item))
                        {
                            Manage_CMD_WyslijWynik_RBR(item);
                            Update_KartaKlienta_RBR(item);
                            Update_StatusZadania(item, StatusZadania.Wysyłka);
                        }
                        break;
                    default:
                        break;
                }
            }
        }

        #region Obsługa przypomnień o terminie płatności


        private int Get_LookupId(SPListItem item, string col)
        {
            return item[col] != null ? new SPFieldLookupValue(item[col].ToString()).LookupId : 0;
        }

        private bool hasPrzypomnienieOTerminiePlatnosci(SPListItem item)
        {
            string col = "colPrzypomnienieOTerminiePlatnos";
            return item[col] != null ? bool.Parse(item[col].ToString()) : false;
        }

        #endregion

        #region Aktualizacja KartyKlienta
        private void Update_KartaKlienta_ZUS(SPListItem item)
        {
            BLL.tabKartyKontrolne.Update_ZUS_Data(item);
        }

        private void Update_KartaKlienta_PD(SPListItem item)
        {
            BLL.tabKartyKontrolne.Update_PD_Data(item);
        }

        private void Update_KartaKlienta_PDS(SPListItem item)
        {
            BLL.tabKartyKontrolne.Update_PDS_Data(item);
        }

        private void Update_KartaKlienta_VAT(SPListItem item)
        {
            BLL.tabKartyKontrolne.Update_VAT_Data(item);
        }
        private void Update_KartaKlienta_RBR(SPListItem item)
        {
            BLL.tabKartyKontrolne.Update_RBR_Data(item);
        }
        #endregion

        private string Get_StatusZadania(SPListItem item)
        {
            return item["enumStatusZadania"] != null ? item["enumStatusZadania"].ToString() : string.Empty;
        }

        private bool isAuditRequest(SPListItem item)
        {
            return item["colAudytDanych"] != null ? (bool)item["colAudytDanych"] : false;
        }


        private void Manage_CMD_WyslijWynik_ProsbaOWyciagBankowy(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {
                //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;

                string nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");
                string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);

                bool KopiaDoNadawcy = false;
                bool KopiaDoBiura = false;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;
                //BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "WBANK_TEMPLATE.Include", out temat, out trescHTML);
                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item.Web, "WBANK_TEMPLATE.Include", out temat, out trescHTML, false);

                string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
                int okresId = Get_LookupId(item, "selOkres");
                string aktualnyMiesiac = BLL.tabOkresy.Get_PoprzedniMiesiacSlownieById(item.Web, okresId, 0);
                if (aktualnyMiesiac != null) aktualnyMiesiac = string.Format(@"({0})", aktualnyMiesiac);
                trescHTML = trescHTML.Replace("___PoprzedniMiesiac___", aktualnyMiesiac);
                trescHTML = trescHTML.Replace("___colInformacjaDlaKlienta___", info);

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
            }
        }

        private void Manage_CMD_WyslijWynik_ProsbaODokumenty(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {
                CreateMessage_ProsbaODokumenty(item, klientId);
            }
        }

        private static void CreateMessage_ProsbaODokumenty(SPListItem item, int klientId)
        {
            //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;
            string nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");
            string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

            string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);

            bool KopiaDoNadawcy = false;
            bool KopiaDoBiura = false;
            string temat = string.Empty;
            string tresc = string.Empty;
            string trescHTML = string.Empty;

            //weź szablon bez stopki
            BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item.Web, "DOK_TEMPLATE.Include", out temat, out trescHTML, false);

            //dodaj nazwę firmy w tytule wiadomości
            temat = BLL.Tools.AddCompanyName(temat, item);

            string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
            trescHTML = trescHTML.Replace("___colInformacjaDlaKlienta___", info);

            DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

            BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
        }


        private void Manage_CMD_WyslijWynik_ZUS(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {

                //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;

                //string nadawca = Get_WlacicielZadania(item);
                //string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

                string nadawca = Get_CurrentUser(item);
                string kopiaDla = Get_WlacicielZadania(item);

                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);

                bool KopiaDoNadawcy = true;
                bool KopiaDoBiura = true;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;

                //wybór szablonu wiadomości

                string zusOpcja = Get_String(item, "colZUS_Opcja");

                if (Get_FlagValue(item, "colZatrudniaPracownikow"))
                {
                    if (GetValue(item, "colZUS_PIT-4R") > 0 || GetValue(item, "colZUS_PIT-8AR") > 0)
                    {
                        switch (zusOpcja)
                        {
                            case "Tylko zdrowotna":
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_Z_PIT_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                                break;
                            default:
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_S_Z_F_PIT_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                                break;
                        }
                    }
                    else
                    {
                        switch (zusOpcja)
                        {
                            case "Tylko zdrowotna":
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_Z_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                                break;
                            default:
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_S_Z_F_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                                break;
                        }
                    }
                }
                else
                {
                    switch (zusOpcja)
                    {
                        case "Tylko zdrowotna":
                            BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_Z_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                            break;
                        default:
                            BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_S_Z_F_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                            break;
                    }
                }

                string lt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_LEADING_TEXT", false);
                string firma = BLL.tabKlienci.Get_NazwaFirmyById(item.Web, klientId);
                lt = lt.Replace("___FIRMA___", firma);
                string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;
                lt = lt.Replace("___OKRES___", okres);
                trescHTML = trescHTML.Replace("___ZUS_LEADING_TEXT___", lt);

                //uzupełnia temat kodem klienta i okresu
                temat = AddSpecyfikacja(item, temat);

                //uzupełnia dane w formatce BR_TEMPLATE
                StringBuilder sb = new StringBuilder(trescHTML);
                sb.Replace("___colZUS_SP_Skladka___", Format_Currency(item, "colZUS_SP_Skladka"));
                sb.Replace("___colZUS_SP_Konto___", item["colZUS_SP_Konto"] != null ? item["colZUS_SP_Konto"].ToString() : string.Empty);
                sb.Replace("___colZUS_TerminPlatnosciSkladek___", Format_Date(item, "colZUS_TerminPlatnosciSkladek"));
                sb.Replace("___colZUS_ZD_Skladka___", Format_Currency(item, "colZUS_ZD_Skladka"));
                sb.Replace("___colZUS_ZD_Konto___", item["colZUS_ZD_Konto"] != null ? item["colZUS_ZD_Konto"].ToString() : string.Empty);
                sb.Replace("___colZUS_FP_Skladka___", Format_Currency(item, "colZUS_FP_Skladka"));
                sb.Replace("___colZUS_FP_Konto___", item["colZUS_FP_Konto"] != null ? item["colZUS_FP_Konto"].ToString() : string.Empty);

                sb.Replace("___colZUS_PIT-4R___", Format_Currency(item, "colZUS_PIT-4R"));
                sb.Replace("___colZUS_PIT-8AR___", Format_Currency(item, "colZUS_PIT-8AR"));

                Klient k = new Klient(item.Web, klientId);

                sb.Replace("___colPIT_Konto___", k.NumerRachunkuPD);

                int okresId = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupId : 0;

                DateTime terminPlatnosciPodatku = BLL.tabOkresy.Get_TerminPlatnosciByOkresId(item.Web, "colPD_TerminPlatnosciPodatku", okresId);
                sb.Replace("___colZUS_TerminPlatnosciPodatku___", terminPlatnosciPodatku.ToShortDateString());

                string info2 = string.Empty;
                string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
                //dodaj informację o z załącznikach w/g ustawionych flag
                if (item["colZUS_PIT-4R_Zalaczony"] != null ? (bool)item["colZUS_PIT-4R_Zalaczony"] : false)
                {
                    if (Get_Flag(item, "colDrukWplaty"))
                    {
                        info2 = info2 + string.Format(templateR, "Druk wpłaty PIT-4R");
                    }
                }
                if (item["colZUS_PIT-8AR_Zalaczony"] != null ? (bool)item["colZUS_PIT-8AR_Zalaczony"] : false)
                {
                    if (Get_Flag(item, "colDrukWplaty"))
                    {
                        info2 = info2 + string.Format(templateR, "Druk wpłaty PIT-8AR");
                    }
                }
                if (item["colZUS_ListaPlac_Zalaczona"] != null ? (bool)item["colZUS_ListaPlac_Zalaczona"] : false)
                {
                    info2 = info2 + string.Format(templateR, "Lista płac");
                }
                if (item["colZUS_Rachunki_Zalaczone"] != null ? (bool)item["colZUS_Rachunki_Zalaczone"] : false)
                {
                    info2 = info2 + string.Format(templateR, "Rachunki");
                }
                if (item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false)
                {
                    info2 = info2 + string.Format(templateR, "Druk(i) wpłaty ZUS");
                }

                if (!string.IsNullOrEmpty(info2))
                {
                    info2 = string.Format(templateH, info2);
                    info = info + "<br>" + info2;
                }


                sb.Replace("___colInformacjaDlaKlienta___", info);

                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                odbiorca = Check_NieWysylacDoKlientaFlag(item, nadawca, odbiorca);

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);


                //reminders
                if (hasPrzypomnienieOTerminiePlatnosci(item))
                {
                    //składki zus
                    DateTime terminPlatnosci = Get_Date(item, "colZUS_TerminPlatnosciSkladek");

                    if (GetValue(item, "colZUS_SP_Skladka") > 0
                        || GetValue(item, "colZUS_ZD_Skladka") > 0
                        || GetValue(item, "colZUS_FP_Skladka") > 0)
                    {
                        //ustaw reminder

                        switch (zusOpcja)
                        {
                            case "Tylko zdrowotna":
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_Z_REMINDER_TEMPLATE.Include", out temat, out trescHTML);
                                break;
                            default:
                                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_S_Z_F_REMINDER_TEMPLATE.Include", out temat, out trescHTML);
                                break;
                        }


                        temat = Update_Data(temat, terminPlatnosci);
                        temat = BLL.Tools.AddCompanyName(temat, item);

                        //leading reminder text
                        string lrt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_LEADING_REMINDER_TEXT", false);
                        lrt = lrt.Replace("___FIRMA___", firma);
                        lrt = lrt.Replace("___OKRES___", okres);
                        trescHTML = trescHTML.Replace("___ZUS_LEADING_REMINDER_TEXT___", lrt);

                        //trailing reminder text
                        string trt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_TRAILING_REMINDER_TEXT", false);
                        trt = trt.Replace("___DATA___", DateTime.Now.ToShortDateString()); //zakłada że wysyłka oryginalnej wiadomości wyjdzie w dniu zlecenia
                        trescHTML = trescHTML.Replace("___ZUS_TRAILING_REMINDER_TEXT___", trt);

                        //aktualizacja danych z tabelki
                        sb = new StringBuilder(trescHTML);
                        sb.Replace("___colZUS_SP_Skladka___", Format_Currency(item, "colZUS_SP_Skladka"));
                        sb.Replace("___colZUS_SP_Konto___", item["colZUS_SP_Konto"] != null ? item["colZUS_SP_Konto"].ToString() : string.Empty);
                        sb.Replace("___colZUS_TerminPlatnosciSkladek___", Format_Date(item, "colZUS_TerminPlatnosciSkladek"));
                        sb.Replace("___colZUS_ZD_Skladka___", Format_Currency(item, "colZUS_ZD_Skladka"));
                        sb.Replace("___colZUS_ZD_Konto___", item["colZUS_ZD_Konto"] != null ? item["colZUS_ZD_Konto"].ToString() : string.Empty);
                        sb.Replace("___colZUS_FP_Skladka___", Format_Currency(item, "colZUS_FP_Skladka"));
                        sb.Replace("___colZUS_FP_Konto___", item["colZUS_FP_Konto"] != null ? item["colZUS_FP_Konto"].ToString() : string.Empty);

                        trescHTML = sb.ToString();


                        planowanaDataNadania = Calc_ReminderTime(item, terminPlatnosci);
                        nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");

                        BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId, Marker.ReminderZUS);
                    }

                    //podatek za pracowników

                    terminPlatnosci = BLL.tabOkresy.Get_TerminPlatnosciByOkresId(item.Web, "colPD_TerminPlatnosciPodatku", okresId);

                    if (GetValue(item, "colZUS_PIT-4R") > 0
                        || GetValue(item, "colZUS_PIT-8AR") > 0)
                    {
                        //ustaw reminder

                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_PIT_REMINDER_TEMPLATE.Include", out temat, out trescHTML);

                        temat = Update_Data(temat, terminPlatnosci);
                        temat = BLL.Tools.AddCompanyName(temat, item);

                        //leading reminder text
                        string lrt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_PIT_LEADING_REMINDER_TEXT", false);
                        lrt = lrt.Replace("___FIRMA___", firma);
                        lrt = lrt.Replace("___OKRES___", okres);
                        trescHTML = trescHTML.Replace("___ZUS_PIT_LEADING_REMINDER_TEXT___", lrt);

                        //trailing reminder text
                        string trt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "ZUS_PIT_TRAILING_REMINDER_TEXT", false);
                        trt = trt.Replace("___DATA___", DateTime.Now.ToShortDateString()); //zakłada że wysyłka oryginalnej wiadomości wyjdzie w dniu zlecenia
                        trescHTML = trescHTML.Replace("___ZUS_PIT_TRAILING_REMINDER_TEXT___", trt);

                        //aktualizacja danych z tabelki
                        sb = new StringBuilder(trescHTML);
                        sb.Replace("___colPIT_Konto___", k.NumerRachunkuPD);
                        sb.Replace("___colZUS_PIT-4R___", Format_Currency(item, "colZUS_PIT-4R"));
                        sb.Replace("___colZUS_PIT-8AR___", Format_Currency(item, "colZUS_PIT-8AR"));
                        sb.Replace("___colZUS_TerminPlatnosciPodatku___", terminPlatnosciPodatku.ToShortDateString());

                        trescHTML = sb.ToString();

                        planowanaDataNadania = Calc_ReminderTime(item, terminPlatnosci);
                        nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");

                        BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId, Marker.ReminderZUS_PIT);
                    }
                }


            }

        }

        private string Format_Date(SPListItem item, string colName)
        {
            DateTime d = DateTime.Parse(Get_String(item, colName));
            return d.ToShortDateString();
        }

        private string Format_Currency(SPListItem item, string colName)
        {
            double n = GetValue(item, colName);

            if (n > 0) return n.ToString("c", new CultureInfo("pl-PL"));
            else return emptyMarker;

        }

        private void Manage_CMD_WyslijWynik_PD(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {
                //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;

                //string nadawca = Get_WlacicielZadania(item);
                //string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

                string nadawca = Get_CurrentUser(item);
                string kopiaDla = Get_WlacicielZadania(item);

                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);

                bool KopiaDoNadawcy = true;
                bool KopiaDoBiura = true;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;

                switch (Get_String(item, "colPD_OcenaWyniku"))
                {
                    case "Dochód":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_DOCHOD_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        //jeżeli wartość do zapłaty = 0 wtdy zastąp tekst formułką i ukryj tabelkę z płatnościami
                        if (GetValue(item, "colPD_WartoscDoZaplaty") == 0)
                            trescHTML = trescHTML.Replace("___NOTIFICATION___", BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_DOCHOD_0_NOTIFICATION", false));
                        else
                            trescHTML = trescHTML.Replace("___NOTIFICATION___", "WARTOŚĆ DO ZAPŁATY");
                        break;
                    case "Strata":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_STRATA_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        break;
                    default:
                        break;
                }

                string lt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_LEADING_TEXT", false);
                string firma = BLL.tabKlienci.Get_NazwaFirmyById(item.Web, klientId);
                lt = lt.Replace("___FIRMA___", firma);

                //opis okresu rozliczeniowego
                string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;

                if (Get_String(item, "enumRozliczeniePD") == "Kwartalnie")
                {
                    okres = "kwartał " + BLL.Tools.Get_KwartalDisplayName(okres);
                }
                else
                {
                    okres = "miesiąc " + okres;
                }
                lt = lt.Replace("___OKRES___", okres);
                trescHTML = trescHTML.Replace("___PD_LEADING_TEXT___", lt);

                //VAT alert
                string va = string.Empty;
                int okresId = Get_LookupId(item, "selOkres");
                int vatZadanieId = BLL.tabZadania.Get_NumerZadaniaVAT(item.Web, klientId, okresId);
                if (vatZadanieId > 0)
                {
                    va = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_VAT_ALERT_TEXT", false);
                }
                trescHTML = trescHTML.Replace("___PD_VAT_ALERT_TEXT___", va);


                //uzupełnia temat kodem klienta i okresu
                temat = AddSpecyfikacja(item, temat);


                //uzupełnia dane w formatce PD_TEMPLATE
                StringBuilder sb = new StringBuilder(trescHTML);

                sb.Replace("___colPD_OcenaWyniku___", Get_String(item, "colPD_OcenaWyniku"));
                sb.Replace("___colPD_WartoscDochodu___", Format_Currency(item, "colPD_WartoscDochodu"));
                sb.Replace("___colPD_WysokoscStraty___", Format_Currency(item, "colPD_WartoscStraty")); //nazwa kolumny rozbieżna
                sb.Replace("___colFormaOpodatkowaniaPD___", Get_String(item, "colFormaOpodatkowaniaPD"));
                sb.Replace("___colPD_WartoscDoZaplaty___", Format_Currency(item, "colPD_WartoscDoZaplaty"));
                sb.Replace("___colPD_Konto___", Get_String(item, "colPD_Konto"));
                sb.Replace("___colPD_TerminPlatnosciPodatku___", Format_Date(item, "colPD_TerminPlatnosciPodatku"));

                string info2 = string.Empty;
                string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
                //dodaj informację o z załącznikach w/g ustawionych flag

                if (Get_String(item, "colPD_OcenaWyniku") == "Dochód"
                && item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false)
                {
                    if (GetValue(item, "colPD_WartoscDoZaplaty") > 0)
                    {
                        info2 = info2 + string.Format(templateR, "Druk wpłaty");
                    }
                }

                if (!string.IsNullOrEmpty(info2))
                {
                    info2 = string.Format(templateH, info2);
                    info = info + "<br>" + info2;
                }

                sb.Replace("___colInformacjaDlaKlienta___", info);

                //ukrywanie zbędnych elementów
                string ocena = Get_String(item, "colPD_OcenaWyniku");
                switch (ocena)
                {
                    case "Dochód":
                        if (GetValue(item, "colPD_WartoscDoZaplaty") > 0)
                            sb.Replace("___Display_T_Platnosc___", string.Empty);
                        break;
                    case "Strata":

                        break;
                    default:

                        break;
                }
                //czyszczenie parametrów
                sb.Replace("___Display_T_Platnosc___", "none");
                sb.Replace("___OpisDochodu_Straty___", string.Empty);
                sb.Replace("___colPD_WartoscDochodu_Straty___", string.Empty);

                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                odbiorca = Check_NieWysylacDoKlientaFlag(item, nadawca, odbiorca);

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);

                //obsługa remindera
                if (hasPrzypomnienieOTerminiePlatnosci(item))
                {
                    DateTime terminPlatnosci = Get_Date(item, "colPD_TerminPlatnosciPodatku");

                    if (Get_String(item, "colPD_OcenaWyniku") == "Dochód")
                    {
                        if (GetValue(item, "colPD_WartoscDoZaplaty") > 0)
                        {
                            //ustaw reminder
                            BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_DOCHOD_REMINDER_TEMPLATE.Include", out temat, out trescHTML);
                            temat = Update_Data(temat, terminPlatnosci);
                            temat = BLL.Tools.AddCompanyName(temat, item);

                            //leading reminder text
                            string lrt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_LEADING_REMINDER_TEXT", false);
                            lrt = lrt.Replace("___FIRMA___", firma);
                            lrt = lrt.Replace("___OKRES___", okres);
                            trescHTML = trescHTML.Replace("___PD_LEADING_REMINDER_TEXT___", lrt);

                            //trailing reminder text
                            string trt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "PD_TRAILING_REMINDER_TEXT", false);
                            trt = trt.Replace("___DATA___", DateTime.Now.ToShortDateString()); //zakłada że wysyłka oryginalnej wiadomości wyjdzie w dniu zlecenia
                            trescHTML = trescHTML.Replace("___PD_TRAILING_REMINDER_TEXT___", trt);

                            //aktualizacja danych z tabelki
                            sb = new StringBuilder(trescHTML);
                            sb.Replace("___colFormaOpodatkowaniaPD___", Get_String(item, "colFormaOpodatkowaniaPD"));
                            sb.Replace("___colPD_WartoscDoZaplaty___", Format_Currency(item, "colPD_WartoscDoZaplaty"));
                            sb.Replace("___colPD_Konto___", Get_String(item, "colPD_Konto"));
                            sb.Replace("___colPD_TerminPlatnosciPodatku___", Format_Date(item, "colPD_TerminPlatnosciPodatku"));

                            trescHTML = sb.ToString();

                            planowanaDataNadania = Calc_ReminderTime(item, terminPlatnosci);
                            nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");

                            BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
                        }
                    }
                }

            }

        }

        private string Get_CurrentUser(SPListItem item)
        {
            string result = item["Editor"] != null ? new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email : string.Empty;

            if (string.IsNullOrEmpty(result))
            {
                //ustaw domyślnie adres biura
                result = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");
            }

            return result;
        }

        private string Update_Data(string temat, DateTime terminPlatnosci)
        {
            return temat.Replace("___DATA___", terminPlatnosci.ToShortDateString());
        }

        /// <summary>
        /// domyślnym nadawcą wiadomości jest bieżący operator a jeżeli go nie ma to biuro
        /// </summary>
        private string Get_WlacicielZadania(SPListItem item)
        {
            string result = string.Empty;
            int operatorId = Get_LookupId(item, "selOperator");
            if (operatorId > 0)
            {
                result = BLL.dicOperatorzy.Get_EmailById(item.Web, operatorId);
            }

            if (string.IsNullOrEmpty(result))
            {
                result = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");
            }

            return result;

        }

        private static string Calc_ReminderSubject(SPListItem item, string kodFormatki, DateTime terminPlatnosci)
        {
            string result = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, kodFormatki, false);
            result = result.Replace("___DATA___", terminPlatnosci.ToShortDateString());
            return result;
        }

        private static DateTime Calc_ReminderTime(SPListItem item, DateTime terminPlatnosci)
        {
            //ustaw datę powiadomienia
            int reminderDateOffset = -1 * int.Parse(BLL.admSetup.GetValue(item.Web, "REMINDER_DATE_OFFSET"));
            if (reminderDateOffset >= 0) reminderDateOffset = -1;
            DateTime reminderDate = terminPlatnosci.AddDays(reminderDateOffset);

            //ustaw godzinę wysyłki powiadomienia
            TimeSpan ts = new TimeSpan(0, 8, 15);
            string reminderTime = BLL.admSetup.GetValue(item.Web, "REMINDER_TIME");
            if (reminderTime.Length == 5) TimeSpan.TryParse(reminderTime, out ts);
            reminderDate = new DateTime(reminderDate.Year, reminderDate.Month, reminderDate.Day, ts.Hours, ts.Minutes, ts.Seconds);
            return reminderDate;
        }

        private void Manage_CMD_WyslijWynik_PDS(SPListItem item)
        {
            //obsługa wysyłki informacji identyczna jak w przypadku PD
            Manage_CMD_WyslijWynik_PD(item);
        }

        private void Manage_CMD_WyslijWynik_VAT(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {

                //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;

                //string nadawca = Get_WlacicielZadania(item);
                //string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

                string nadawca = Get_CurrentUser(item);
                string kopiaDla = Get_WlacicielZadania(item);


                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);
                bool KopiaDoNadawcy = true;
                bool KopiaDoBiura = true;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;

                bool IsBWAllowed = false; //Czy informacja o blankicie wpłaty może być załączona

                switch (item["colVAT_Decyzja"] != null ? item["colVAT_Decyzja"].ToString() : string.Empty)
                {
                    case "Do zapłaty":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_DO_ZAPLATY_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        IsBWAllowed = true;
                        break;
                    case "Do przeniesienia":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_DO_PRZENIESIENIA_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        break;
                    case "Do zwrotu":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_DO_ZWROTU_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        break;
                    case "Do przeniesienia i do zwrotu":
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_DO_PRZENIESIENIA_ZWROTU_TEMPLATE.Include", out temat, out trescHTML, nadawca);
                        break;
                    default:
                        break;
                }

                string lt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_LEADING_TEXT", false);
                string firma = BLL.tabKlienci.Get_NazwaFirmyById(item.Web, klientId);
                lt = lt.Replace("___FIRMA___", firma);

                //zdefiniuj opis bieżącego okresu
                string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;

                if (Get_String(item, "enumRozliczenieVAT") == "Kwartalnie")
                {
                    okres = "kwartał " + BLL.Tools.Get_KwartalDisplayName(okres);
                }
                else
                {
                    okres = "miesiąc " + okres;
                }

                lt = lt.Replace("___OKRES___", okres);
                trescHTML = trescHTML.Replace("___VAT_LEADING_TEXT___", lt);


                //uzupełnia temat kodem klienta i okresu
                temat = AddSpecyfikacja(item, temat);

                //uzupełnia dane w formatce VAT_TEMPLATE 
                StringBuilder sb = new StringBuilder(trescHTML);
                sb.Replace("___colVAT_Decyzja___", item["colVAT_Decyzja"] != null ? item["colVAT_Decyzja"].ToString() : string.Empty);
                sb.Replace("___colVAT_TerminZwrotuPodatku___", item["colVAT_TerminZwrotuPodatku"] != null ? item["colVAT_TerminZwrotuPodatku"].ToString() : "?");
                sb.Replace("___colVAT_WartoscNadwyzkiZaPoprzedniMiesiac___", Format_Currency(item, "colVAT_WartoscNadwyzkiZaPoprzedniMiesiac"));
                sb.Replace("___colVAT_WartoscDoZwrotu___", Format_Currency(item, "colVAT_WartoscDoZwrotu"));
                sb.Replace("___colVAT_WartoscDoPrzeniesienia___", Format_Currency(item, "colVAT_WartoscDoPrzeniesienia"));
                sb.Replace("___colFormaOpodatkowaniaVAT___", item["colFormaOpodatkowaniaVAT"] != null ? item["colFormaOpodatkowaniaVAT"].ToString() : string.Empty);
                sb.Replace("___colVAT_WartoscDoZaplaty___", Format_Currency(item, "colVAT_WartoscDoZaplaty"));
                sb.Replace("___colVAT_Konto___", item["colVAT_Konto"] != null ? item["colVAT_Konto"].ToString() : string.Empty);
                sb.Replace("___colVAT_TerminPlatnosciPodatku___", item["colVAT_TerminPlatnosciPodatku"] != null ? DateTime.Parse(item["colVAT_TerminPlatnosciPodatku"].ToString()).ToShortDateString() : string.Empty);

                string info2 = string.Empty;
                string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
                //dodaj informację o z załącznikach w/g ustawionych flag
                //if (item["colVAT_VAT-UE_Zalaczony"] != null ? (bool)item["colVAT_VAT-UE_Zalaczony"] : false)
                //{
                //    info2 = info2 + string.Format(templateR, "VAT-UE");
                //}
                //if (item["colVAT_VAT_x002d_27_Zalaczony0"] != null ? (bool)item["colVAT_VAT_x002d_27_Zalaczony0"] : false)
                //{
                //    info2 = info2 + string.Format(templateR, "VAT-27");
                //}

                if (item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false &&
                    IsBWAllowed) //dodawaj informację o załącznikach tylko w przypadku płatności VAT
                {
                    info2 = info2 + string.Format(templateR, "Druk wpłaty");
                }

                if (!string.IsNullOrEmpty(info2))
                {
                    info2 = string.Format(templateH, info2);
                    info = info + "<br>" + info2;
                }


                sb.Replace("___colInformacjaDlaKlienta___", info);

                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                odbiorca = Check_NieWysylacDoKlientaFlag(item, nadawca, odbiorca);

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);

                //obsługa remindera
                if (hasPrzypomnienieOTerminiePlatnosci(item))
                {
                    DateTime terminPlatnosci = Get_Date(item, "colVAT_TerminPlatnosciPodatku");

                    if (Get_String(item, "colVAT_Decyzja") == "Do zapłaty")
                    {
                        if (GetValue(item, "colVAT_WartoscDoZaplaty") > 0)
                        {
                            //ustaw reminder
                            BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_DO_ZAPLATY_REMINDER_TEMPLATE.Include", out temat, out trescHTML);
                            temat = Update_Data(temat, terminPlatnosci);
                            temat = BLL.Tools.AddCompanyName(temat, item);

                            //leading reminder text
                            string lrt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_LEADING_REMINDER_TEXT", false);
                            lrt = lrt.Replace("___FIRMA___", firma);
                            lrt = lrt.Replace("___OKRES___", okres);
                            trescHTML = trescHTML.Replace("___VAT_LEADING_REMINDER_TEXT___", lrt);

                            //trailing reminder text
                            string trt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "VAT_TRAILING_REMINDER_TEXT", false);
                            trt = trt.Replace("___DATA___", DateTime.Now.ToShortDateString()); //zakłada że wysyłka oryginalnej wiadomości wyjdzie w dniu zlecenia
                            trescHTML = trescHTML.Replace("___VAT_TRAILING_REMINDER_TEXT___", trt);

                            //aktualizacja danych z tabelki
                            sb = new StringBuilder(trescHTML);
                            sb.Replace("___colFormaOpodatkowaniaVAT___", item["colFormaOpodatkowaniaVAT"] != null ? item["colFormaOpodatkowaniaVAT"].ToString() : string.Empty);
                            sb.Replace("___colVAT_WartoscDoZaplaty___", Format_Currency(item, "colVAT_WartoscDoZaplaty"));
                            sb.Replace("___colVAT_Konto___", item["colVAT_Konto"] != null ? item["colVAT_Konto"].ToString() : string.Empty);
                            sb.Replace("___colVAT_TerminPlatnosciPodatku___", item["colVAT_TerminPlatnosciPodatku"] != null ? DateTime.Parse(item["colVAT_TerminPlatnosciPodatku"].ToString()).ToShortDateString() : string.Empty);

                            trescHTML = sb.ToString();

                            planowanaDataNadania = Calc_ReminderTime(item, terminPlatnosci);
                            nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");

                            BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
                        }
                    }
                }
            }
        }

        private void Manage_CMD_WyslijWynik_RBR(SPListItem item)
        {
            string cmd = GetCommand(item);
            int klientId = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupId : 0;

            if (klientId > 0
                && cmd == ZATWIERDZ)
            {
                //string nadawca = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User.Email;

                //string nadawca = Get_WlacicielZadania(item);
                //string kopiaDla = Get_KopiaDlaEdytora(item, nadawca);

                string nadawca = Get_CurrentUser(item);
                string kopiaDla = Get_WlacicielZadania(item);

                string odbiorca = BLL.tabKlienci.Get_EmailById(item.Web, new SPFieldLookupValue(item["selKlient"].ToString()).LookupId);


                bool KopiaDoNadawcy = false;
                bool KopiaDoBiura = true;
                string temat = string.Empty;
                string tresc = string.Empty;
                string trescHTML = string.Empty;

                BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "RBR_TEMPLATE.Include", out temat, out trescHTML, nadawca);

                string lt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "RBR_LEADING_TEXT", false);
                string firma = BLL.tabKlienci.Get_NazwaFirmyById(item.Web, klientId);
                lt = lt.Replace("___FIRMA___", firma);
                string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;
                lt = lt.Replace("___OKRES___", okres);
                trescHTML = trescHTML.Replace("___RBR_LEADING_TEXT___", lt);

                //uzupełnia temat kodem klienta i okresu
                temat = AddSpecyfikacja(item, temat);

                //uzupełnia dane w formatce BR_TEMPLATE
                StringBuilder sb = new StringBuilder(trescHTML);
                sb.Replace("___colBR_NumerFaktury___", item["colBR_NumerFaktury"] != null ? item["colBR_NumerFaktury"].ToString() : string.Empty);
                sb.Replace("___colBR_DataWystawienia___", Format_Date(item, "colBR_DataWystawieniaFaktury"));
                sb.Replace("___colBR_WartoscDoZaplaty___", BLL.Tools.Format_Currency(item["colBR_WartoscDoZaplaty"] != null ? double.Parse(item["colBR_WartoscDoZaplaty"].ToString()) : 0));
                sb.Replace("___colBR_Konto___", item["colBR_Konto"] != null ? item["colBR_Konto"].ToString() : string.Empty);
                sb.Replace("___colBR_TerminPlatnosci___", Format_Date(item, "colBR_TerminPlatnosci"));

                string info2 = string.Empty;
                string info = item["colInformacjaDlaKlienta"] != null ? item["colInformacjaDlaKlienta"].ToString() : string.Empty;
                //dodaj informację o z załącznikach w/g ustawionych flag
                if (item["colBR_FakturaZalaczona"] != null ? (bool)item["colBR_FakturaZalaczona"] : false)
                {
                    info2 = info2 + string.Format(templateR, "Faktura za obsługę księgową");
                }
                if (item["colDrukWplaty"] != null ? (bool)item["colDrukWplaty"] : false)
                {
                    info2 = info2 + string.Format(templateR, "Druk wpłaty");
                }

                if (!string.IsNullOrEmpty(info2))
                {
                    info2 = string.Format(templateH, info2);
                    info = info + info2;
                }

                sb.Replace("___colInformacjaDlaKlienta___", info);

                trescHTML = sb.ToString();

                DateTime planowanaDataNadania = item["colTerminWyslaniaInformacji"] != null ? DateTime.Parse(item["colTerminWyslaniaInformacji"].ToString()) : new DateTime();

                BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);

                //obsługa remindera
                if (hasPrzypomnienieOTerminiePlatnosci(item))
                {
                    DateTime terminPlatnosci = Get_Date(item, "colBR_TerminPlatnosci");


                    if (GetValue(item, "colBR_WartoscDoZaplaty") > 0)
                    {
                        //ustaw reminder
                        BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "RBR_REMINDER_TEMPLATE.Include", out temat, out trescHTML);
                        temat = Update_Data(temat, terminPlatnosci);
                        temat = BLL.Tools.AddCompanyName(temat, item);

                        //leading reminder text
                        string lrt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "RBR_LEADING_REMINDER_TEXT", false);
                        lrt = lrt.Replace("___FIRMA___", firma);
                        lrt = lrt.Replace("___OKRES___", okres);
                        trescHTML = trescHTML.Replace("___RBR_LEADING_REMINDER_TEXT___", lrt);

                        //trailing reminder text
                        string trt = BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item, "RBR_TRAILING_REMINDER_TEXT", false);
                        trt = trt.Replace("___DATA___", DateTime.Now.ToShortDateString()); //zakłada że wysyłka oryginalnej wiadomości wyjdzie w dniu zlecenia
                        trescHTML = trescHTML.Replace("___RBR_TRAILING_REMINDER_TEXT___", trt);

                        //aktualizacja danych z tabelki
                        sb = new StringBuilder(trescHTML);
                        sb.Replace("___colBR_NumerFaktury___", item["colBR_NumerFaktury"] != null ? item["colBR_NumerFaktury"].ToString() : string.Empty);
                        sb.Replace("___colBR_DataWystawienia___", Format_Date(item, "colBR_DataWystawieniaFaktury"));
                        sb.Replace("___colBR_WartoscDoZaplaty___", Format_Currency(item, "colBR_WartoscDoZaplaty"));
                        sb.Replace("___colBR_Konto___", item["colBR_Konto"] != null ? item["colBR_Konto"].ToString() : string.Empty);
                        sb.Replace("___colBR_TerminPlatnosci___", Format_Date(item, "colBR_TerminPlatnosci"));

                        trescHTML = sb.ToString();

                        planowanaDataNadania = Calc_ReminderTime(item, terminPlatnosci);
                        nadawca = BLL.admSetup.GetValue(item.Web, "EMAIL_BIURA");

                        BLL.tabWiadomosci.AddNew(item.Web, item, nadawca, odbiorca, kopiaDla, KopiaDoNadawcy, KopiaDoBiura, temat, tresc, trescHTML, planowanaDataNadania, item.ID, klientId);
                    }

                }
            }
        }

        private static string Get_KopiaDlaEdytora(SPListItem item, string nadawca)
        {
            string result = string.Empty;

            SPUser user = new SPFieldUserValue(item.Web, item["Editor"].ToString()).User;
            //sprawdź przypisanie adresu na liście operatorów
            int operatorId = BLL.dicOperatorzy.Get_OperatorIdByLoginName(item.Web, user.LoginName);
            if (operatorId > 0)
            {
                result = BLL.dicOperatorzy.Get_EmailById(item.Web, operatorId);
            }

            if (result == nadawca)
            {
                //jeżeli operator nie ma przypisanego adresu mailowego lub pokrywa się z adresem nadawcy
                return string.Empty;
            }

            return result;
        }


        private bool isValidated_ZUS(SPListItem item)
        {
            Set_ValidationFlag(item, false);

            //oczyść dane w zależności od wybranej Decyzji
            string opcja = item["colZUS_Opcja"] != null ? item["colZUS_Opcja"].ToString() : string.Empty;
            if (string.IsNullOrEmpty(opcja))
            {
                return false;
            }

            switch (opcja)
            {
                case "Tylko zdrowotna":
                    ClearValue(item, "colZUS_SP_Skladka");
                    ClearValue(item, "colZUS_FP_Skladka");
                    if (GetValue(item, "colZUS_ZD_Skladka") >= 0)
                    {
                        bool foundError = false;

                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_SP_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS SP");
                            foundError = true;
                        }
                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_ZD_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS ZD");
                            foundError = true;
                        }
                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_FP_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS FP");
                            foundError = true;
                        }

                        if (foundError)
                        {
                            Set_ValidationFlag(item, true);
                            return false;
                        }

                        return true;
                    }
                    break;
                default:
                    if (GetValue(item, "colZUS_SP_Skladka") >= 0
                        && GetValue(item, "colZUS_ZD_Skladka") >= 0
                        && GetValue(item, "colZUS_FP_Skladka") >= 0)
                    {
                        bool foundError = false;

                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_SP_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS SP");
                            foundError = true;
                        }
                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_ZD_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS ZD");
                            foundError = true;
                        }
                        if (string.IsNullOrEmpty(Get_String(item, "colZUS_FP_Konto")))
                        {
                            Add_Comment(item, "brak numeru konta ZUS FP");
                            foundError = true;
                        }

                        if (foundError)
                        {
                            Set_ValidationFlag(item, true);
                            return false;
                        }

                        return true;
                    }
                    break;
            }

            Set_ValidationFlag(item, true);
            return false;
        }

        private bool isValidated_PD(SPListItem item)
        {
            Set_ValidationFlag(item, false);

            //oczyść dane w zależności od wybranej Decyzji
            string ocena = Get_String(item, "colPD_OcenaWyniku");
            if (string.IsNullOrEmpty(ocena))
            {
                Add_Comment(item, "brak oceny wyniku");
                Set_ValidationFlag(item, true);
                return false;
            }

            switch (ocena)
            {
                case "Dochód":
                    ClearValue(item, "colPD_WartoscStraty");

                    if (GetValue(item, "colPD_WartoscDoZaplaty") >= 0
                        && GetValue(item, "colPD_WartoscDochodu") >= 0)
                        if (!string.IsNullOrEmpty(Get_String(item, "colPD_Konto"))) return true;
                        else
                        {
                            Add_Comment(item, "brak numeru konta");
                            Set_ValidationFlag(item, true);
                        }
                    break;
                case "Strata":
                    ClearValue(item, "colPD_WartoscDochodu");
                    ClearValue(item, "colPD_WartoscDoZaplaty");

                    if (GetValue(item, "colPD_WartoscStraty") > 0) return true;
                    else
                    {
                        Add_Comment(item, "wartość straty musi być większa niż 0");
                        Set_ValidationFlag(item, true);
                    }
                    break;
                default:
                    Add_Comment(item, "niedozwolona wartość pola ocena wyniku");
                    break;
            }

            Set_ValidationFlag(item, true);
            return false;
        }

        private void Set_ValidationFlag(SPListItem item, bool flag)
        {
            BLL.tabZadania.Set_ValidationFlag(item, flag);
        }

        private bool isValidated_PDS(SPListItem item)
        {
            //wszystkie warunki dla PD powinny być spełnione dla PDS
            return isValidated_PD(item);
        }

        private bool isValidated_VAT(SPListItem item)
        {
            Set_ValidationFlag(item, false);

            //oczyść dane w zależności od wybranej Decyzji
            string decyzja = item["colVAT_Decyzja"] != null ? item["colVAT_Decyzja"].ToString() : string.Empty;
            if (string.IsNullOrEmpty(decyzja))
            {
                return false;
            }

            switch (decyzja)
            {
                case "Do zapłaty":
                    //ClearValue(item, "colVAT_WartoscDoZaplaty");
                    ClearValue(item, "colVAT_WartoscDoPrzeniesienia");
                    ClearValue(item, "colVAT_WartoscDoZwrotu");

                    if (GetValue(item, "colVAT_WartoscDoZaplaty") >= 0)
                        if (!string.IsNullOrEmpty(Get_String(item, "colVAT_Konto"))) return true;
                        else Add_Comment(item, "brak numeru konta");
                    else
                    {
                        Add_Comment(item, "Wartość do zapłaty nie może być ujemna");
                    }
                    break;
                case "Do przeniesienia":
                    ClearValue(item, "colVAT_WartoscDoZaplaty");
                    //ClearValue(item, "colVAT_WartoscDoPrzeniesienia");
                    ClearValue(item, "colVAT_WartoscDoZwrotu");

                    if (GetValue(item, "colVAT_WartoscDoPrzeniesienia") >= 0) return true;
                    else
                    {
                        Add_Comment(item, "Wartość do przeniesienia nie może być ujemna");
                    }
                    break;
                case "Do zwrotu":
                    ClearValue(item, "colVAT_WartoscDoZaplaty");
                    ClearValue(item, "colVAT_WartoscDoPrzeniesienia");
                    //ClearValue(item, "colVAT_WartoscDoZwrotu");

                    if (GetValue(item, "colVAT_WartoscDoZwrotu") >= 0) return true;
                    else
                    {
                        Add_Comment(item, "Wartość do zwrotu nie może być ujemna");
                    }
                    break;
                case "Do przeniesienia i do zwrotu":
                    ClearValue(item, "colVAT_WartoscDoZaplaty");
                    //ClearValue(item, "colVAT_WartoscDoPrzeniesienia");
                    //ClearValue(item, "colVAT_WartoscDoZwrotu");

                    if (GetValue(item, "colVAT_WartoscDoPrzeniesienia") >= 0
                        && GetValue(item, "colVAT_WartoscDoZwrotu") >= 0) return true;
                    else
                    {
                        Add_Comment(item, "Niedozwolone wartości ujemne");
                    }
                    break;
                default:
                    break;
            }

            Set_ValidationFlag(item, true);
            return false;
        }

        private void Add_Comment(SPListItem item, string comment)
        {
            string uwagi = Get_String(item, "colUwagi");
            uwagi = uwagi + "\n" + DateTime.Now.ToString() + "\n" + comment;
            item["colUwagi"] = uwagi.Trim();
            item.SystemUpdate();
        }

        private string Get_String(SPListItem item, string colName)
        {
            return item[colName] != null ? item[colName].ToString() : string.Empty;
        }

        private double GetValue(SPListItem item, string colName)
        {
            if (item[colName] != null)
            {
                return double.Parse(item[colName].ToString());
            }
            else
            {
                //jeżeli pusta wartość to wpisz 0
                item[colName] = 0;
                item.SystemUpdate();
                return 0;
            }
        }

        private void ClearValue(SPListItem item, string colName)
        {
            if (item[colName] != null)
            {
                item[colName] = string.Empty;
                item.SystemUpdate();
            }
        }

        private bool isValidated_RBR(SPListItem item)
        {
            Set_ValidationFlag(item, false);

            StringBuilder sb = new StringBuilder();

            bool foundErrors = false;
            if (string.IsNullOrEmpty(Get_String(item, "colBR_NumerFaktury")))
            {
                foundErrors = true;
                sb.AppendLine(@"brak numeru faktury");
            }

            if (GetValue(item, "colBR_WartoscDoZaplaty") <= 0)
            {
                foundErrors = true;
                sb.AppendLine(@"nieprawidłowa wartość do zapłaty");
            }
            if (item["colBR_FakturaZalaczona"] != null ? (bool)item["colBR_FakturaZalaczona"] : false)
            {
                if (item.Attachments.Count == 0)
                {
                    foundErrors = true;
                    sb.AppendLine(@"brak załącznika");
                }
            }

            if (!foundErrors) return true;
            else
            {
                Add_Comment(item, sb.ToString());
                Set_ValidationFlag(item, true);
                return false;
            }

        }

        #endregion

        #region Helpers
        private bool Get_FlagValue(SPListItem item, string col)
        {
            return item[col] != null ? bool.Parse(item[col].ToString()) : false;
        }

        private DateTime Get_Date(SPListItem item, string col)
        {
            return item[col] != null ? DateTime.Parse(item[col].ToString()) : new DateTime();
        }
        private static string Check_NieWysylacDoKlientaFlag(SPListItem item, string nadawca, string odbiorca)
        {
            bool czyNieWysylacDoKlienta = item["colNieWysylajDoKlienta"] != null ? (bool)item["colNieWysylajDoKlienta"] : false;
            if (czyNieWysylacDoKlienta)
            {
                odbiorca = nadawca;
            }
            return odbiorca;
        }
        private string AddSpecyfikacja(SPListItem item, string temat)
        {
            string okres = item["selOkres"] != null ? new SPFieldLookupValue(item["selOkres"].ToString()).LookupValue : string.Empty;
            string klient = item["selKlient"] != null ? new SPFieldLookupValue(item["selKlient"].ToString()).LookupValue : string.Empty;

            if (!string.IsNullOrEmpty(okres)) temat = temat + " " + okres;
            if (!string.IsNullOrEmpty(klient)) temat = temat + " " + klient;

            return temat;
        }
        private string AddSygnatura(string temat, SPListItem item)
        {
            if (item != null)
            {
                return string.Format("{0} [#{1}]", temat, item.ID.ToString());
            }
            return temat;
        }

        private void Update_StatusZadania(SPListItem item, StatusZadania statusZadania)
        {
            item["enumStatusZadania"] = statusZadania.ToString();
            item.SystemUpdate();
        }
        private static void UsunPodobneZalaczniki(SPListItem item, string targetFileNameLeading)
        {
            if (item.Attachments.Count > 0)
            {

                System.Collections.Generic.List<string> foundNames = new System.Collections.Generic.List<string>();

                foreach (string attName in item.Attachments)
                {
                    if (attName.StartsWith(targetFileNameLeading))
                    {
                        foundNames.Add(attName);
                    }
                }

                if (foundNames.Count > 0)
                {
                    foreach (string attName in foundNames)
                    {
                        item.Attachments.Delete(attName);
                    }
                    item.SystemUpdate();
                }

            }
        }
        private void Validate(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;

            string ct = properties.AfterProperties["ContentType"] != null ? properties.AfterProperties["ContentType"].ToString() : string.Empty;
            int klientId = properties.AfterProperties["selKlient"] != null ? new SPFieldLookupValue(properties.AfterProperties["selKlient"].ToString()).LookupId : 0;
            int okresId = properties.AfterProperties["selOkres"] != null ? new SPFieldLookupValue(properties.AfterProperties["selOkres"].ToString()).LookupId : 0;

            if (!string.IsNullOrEmpty(ct)
                && klientId > 0
                && okresId > 0)
            {
                string key = tabZadania.Define_KEY(ct, klientId, okresId);
                using (SPWeb web = properties.Web)
                {
                    properties.Cancel = !tabZadania.Check_KEY_IsAllowed(key, web, properties.ListItemId);
                    properties.ErrorMessage = "Zdublowany klucz zadania";
                }
            }

            this.EventFiringEnabled = true;
        }

        private void ResetCommand(SPListItem item, bool clearInformacjaDlaKlienta)
        {
            item["cmdFormatka"] = string.Empty; //czyszczenie komendy

            if (clearInformacjaDlaKlienta
                && item["colInformacjaDlaKlienta"] != null)
            {
                item["colInformacjaDlaKlienta"] = string.Empty;
            }
            item.SystemUpdate();

        }

        private string GetCommand(SPListItem item)
        {
            return item["cmdFormatka"] != null ? item["cmdFormatka"].ToString() : string.Empty;
        }

        private static void Assign_OperatorBasedOnProcedura(SPListItem item, SPWeb web, int procId)
        {
            if (procId > 0 && item["selOperator"] == null)
            {
                int operatorId = BLL.tabProcedury.Get_OperatorById(web, procId);
                if (operatorId > 0)
                {

                    item["selOperator"] = operatorId;
                    item.SystemUpdate();

                }
            }
        }

        private static void Assign_TerminRealizacjiBasedOnProcedura(SPListItem item, SPWeb web, int procId)
        {
            if (procId > 0 && (item["colTerminRealizacji"] == null || (DateTime)item["colTerminRealizacji"] != new DateTime()))
            {

                int termin = BLL.tabProcedury.Get_TerminRealizacjiOfsetById(web, procId);
                if (termin > 0)
                {
                    item["colTerminRealizacji"] = DateTime.Today.AddDays(termin);
                    item.SystemUpdate();
                }
            }
        }

        private static int Assign_ProceduraBasedOnTitle(SPListItem item, SPWeb web)
        {
            int procId = item["selProcedura"] != null ? new SPFieldLookupValue(item["selProcedura"].ToString()).LookupId : 0;
            if (procId == 0)
            {
                procId = BLL.tabProcedury.Ensure(web, item.Title);
                item["selProcedura"] = procId;
                item.SystemUpdate();
            }
            return procId;
        }
        private bool Get_Flag(SPListItem item, string col)
        {
            return item[col] != null ? bool.Parse(item[col].ToString()) : false;
        }


        #endregion

    }
}
