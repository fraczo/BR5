using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL
{
    public class tabKartyKontrolne
    {
        const string targetList = "Karty kontrolne";
        const string colPOTWIERDZENIE_ODBIORU_DOKUMENTOW = "colPotwierdzenieOdbioruDokumento";

        public static void Create_KartaKontrolna(SPWeb web, int klientId, int okresId)
        {
            string KEY = Create_KEY(klientId, okresId);
            int formId = Get_KartaKontrolnaId(web,klientId, okresId, KEY);
        }

        public static void Update_PD_Data(Microsoft.SharePoint.SPListItem item)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            Copy_PDFields(item, form);
            form.SystemUpdate();
        }

        private static void Copy_PDFields(Microsoft.SharePoint.SPListItem item, SPListItem form)
        {
            Copy_Field(item, form, "colPotwierdzenieOdbioruDokumento");
            Copy_Field(item, form, "colFormaOpodatkowaniaPD");
            Copy_Field(item, form, "enumRozliczeniePD");
            Copy_Field(item, form, "colPD_OcenaWyniku");
            Copy_Field(item, form, "colPD_WartoscDochodu");
            Copy_Field(item, form, "colPD_WartoscDoZaplaty");
            //Copy_Field(item, form, "colPD_WartoscStraty");
            Copy_Field(item, form, "colPD_WartoscStraty", -1);

            Copy_Field(item, "selOperator", form, "selOperator_PD");
            Copy_Field(item, "colUwagi", form, "Uwagi_PD");

            Copy_Field(item, "colNieWysylajDoKlienta", form, "_NieWysylajDoKlienta_PD");

            Copy_Id(item, form, "_ZadanieID_PD");
        }

        public static void Update_PDS_Data(SPListItem item)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            Copy_PDFields(item, form);
            Copy_Field(item, form, "colIloscDokWBPKN");
            Copy_Field(item, form, "colIloscFaktur");
            Copy_Field(item, form, "colKosztyNKUP");
            Copy_Field(item, form, "colKosztyNKUP_WynWyl");
            Copy_Field(item, form, "colKosztyNKUP_ZUSPlatWyl");
            Copy_Field(item, form, "colKosztyNKUP_FakWyl");
            Copy_Field(item, form, "colKosztyNKUP_PozostaleKoszty");
            Copy_Field(item, form, "colKosztyWS");
            Copy_Field(item, form, "colKosztyWS_WynWlaczone");
            Copy_Field(item, form, "colKosztyWS_ZUSPlatWlaczone");
            Copy_Field(item, form, "colKosztyWS_FakWlaczone");
            Copy_Field(item, form, "colPrzychodyNP");
            Copy_Field(item, form, "colPrzychodyNP_DywidendySpO");
            Copy_Field(item, form, "colPrzychodyNP_Inne");
            Copy_Field(item, form, "colPrzychodyZwolnione");
            Copy_Field(item, form, "colWplaconaSZ");
            Copy_Field(item, form, "colZyskStrataNetto");

            Copy_Field(item, form, "colStrataDoOdliczenia");
            Copy_Field(item, form, "colStronaWn");
            Copy_Field(item, form, "colStronaMa");
            Copy_Field(item, form, "colStronaWn-StronaMa");
            Copy_Field(item, form, "colWplaconeZaliczkiOdPoczatkuRoku");
            Copy_Field(item, form, "colIleDoDoplaty");

            BLL.Models.Klient k = new Models.Klient(item.Web, Get_LookupId(item, "selKlient"));
            form["enumFormaPrawna"] = k.FormaPrawna;

            form.SystemUpdate();
        }

        public static void Update_VAT_Data(SPListItem item)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            Copy_Field(item, form, "colFormaOpodatkowaniaVAT");
            Copy_Field(item, form, "enumRozliczenieVAT");
            Copy_Field(item, form, "colVAT_WartoscNadwyzkiZaPoprzedn");
            Copy_Field(item, form, "colVAT_Decyzja");
            Copy_Field(item, form, "colVAT_WartoscDoZaplaty");
            Copy_Field(item, form, "colVAT_WartoscDoPrzeniesienia");
            Copy_Field(item, form, "colVAT_WartoscDoZwrotu");
            Copy_Field(item, form, "colVAT_eDeklaracja");
            Copy_Field(item, form, "colVAT_VAT-UE_Zalaczony");
            Copy_Field(item, form, "colVAT_VAT-27_Zalaczony");

            Copy_Field(item, "selOperator", form, "selOperator_VAT");
            Copy_Field(item, "colUwagi", form, "Uwagi_VAT");

            Copy_Field(item, "colNieWysylajDoKlienta", form, "_NieWysylajDoKlienta_VAT");

            Copy_Id(item, form, "_ZadanieID_VAT");
            
            form.SystemUpdate();
        }

        public static void Update_ZUS_Data(SPListItem item)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            Copy_Field(item, form, "colFormaOpodakowania_ZUS");
            Copy_Field(item, form, "colZUS_SP_Skladka");
            Copy_Field(item, form, "colZUS_ZD_Skladka");
            Copy_Field(item, form, "colZUS_FP_Skladka");
            Copy_Field(item, form, "colZatrudniaPracownikow");
            Copy_Field(item, form, "colZUS_PIT-4R_Zalaczony");
            Copy_Field(item, form, "colZUS_PIT-4R");
            Copy_Field(item, form, "colVAT_eDeklaracja");
            Copy_Field(item, form, "colZUS_PIT-8AR_Zalaczony");
            Copy_Field(item, form, "colZUS_PIT-8AR");
            Copy_Field(item, form, "colZUS_ListaPlac_Zalaczona");
            Copy_Field(item, form, "colZUS_Rachunki_Zalaczone");

            Copy_Field(item, "selOperator", form, "selOperator_ZUS");
            Copy_Field(item, "colUwagi", form, "Uwagi_ZUS");
            Copy_Field(item, form, "colZUS_Rachunki_Zalaczone");
            Copy_Field(item, form, "colUwagiKadrowe");

            BLL.Models.Klient k = new Models.Klient(item.Web, Get_LookupId(item, "selKlient"));
            form["colDataRozpoczeciaDzialalnosci"] = k.DataRozpoczeciaDzialalnosci;

            Copy_Field(item, "colNieWysylajDoKlienta", form, "_NieWysylajDoKlienta_ZUS");

            Copy_Id(item, form, "_ZadanieID_ZUS");

            form.SystemUpdate();
        }

        public static void Update_RBR_Data(SPListItem item)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            Copy_Field(item, form, "colBR_DataWystawieniaFaktury");
            Copy_Field(item, form, "colBR_NumerFaktury");
            Copy_Field(item, form, "colBR_WartoscDoZaplaty");
            Copy_Field(item, form, "colBR_TerminPlatnosci");
            Copy_Field(item, form, "colBR_FakturaZalaczona");

            Copy_Id(item, form, "_ZadanieID_RBR");

            form.SystemUpdate();
        }

        public static void Update_PD_DataWysylki(SPListItem item, DateTime dateTime)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            form["colPD_DataWylaniaInformacji"] = dateTime;
            form.SystemUpdate();
        }

        public static void Update_VAT_DataWysylki(SPListItem item, DateTime dateTime)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            form["colVAT_DataWyslaniaInformacji"] = dateTime;
            form.SystemUpdate();
        }

        public static void Update_ZUS_DataWysylki(SPListItem item, DateTime dateTime)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            form["colZUS_DataWyslaniaInformacji"] = dateTime;
            form.SystemUpdate();
        }

        public static void Update_RBR_DataWysylki(SPListItem item, DateTime dateTime)
        {
            string KEY = Create_KEY(item);
            int formId = Get_KartaKontrolnaId(item, KEY);

            SPListItem form = Get_KartaKontrolnaById(item.Web, formId);
            form["colBR_DataPrzekazania"] = dateTime;
            form.SystemUpdate();
        }


        private static SPListItem Get_KartaKontrolnaById(SPWeb web, int formId)
        {
            SPList list = web.Lists.TryGetList(targetList);
            return list.Items.GetItemById(formId);
        }

        private static int Get_KartaKontrolnaId(SPListItem task, string KEY)
        {
            SPList list = task.Web.Lists.TryGetList(targetList);
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["KEY"].ToString() == KEY)
                .FirstOrDefault();
            if (item != null)
            {
                return item.ID;
            }
            else
            {
                SPListItem newItem = list.AddItem();
                newItem["KEY"] = KEY;
                newItem["selKlient"] = Get_LookupId(task, "selKlient");
                newItem["selOkres"] = Get_LookupId(task, "selOkres");

                BLL.Models.Klient k = new Models.Klient(task.Web, Get_LookupId(task, "selKlient"));

                newItem["enumRozliczeniePD"] = k.RozliczeniePD;
                newItem["enumRozliczenieVAT"] = k.RozliczenieVAT;
                newItem["colFormaOpodatkowaniaPD"] = k.FormaOpodatkowaniaPD;
                newItem["colFormaOpodatkowaniaVAT"] = k.FormaOpodatkowaniaVAT;
                newItem["colFormaOpodakowania_ZUS"] = k.FormaOpodatkowaniaZUS;

                //ustaw CT
                if (k.TypKlienta == "KSH") newItem["ContentType"] = "Karta kontrolna KSH";
                else newItem["ContentType"] = "Karta kontrolna KPiR";

                newItem.SystemUpdate();

                return newItem.ID;
            }

        }

        private static int Get_KartaKontrolnaId(SPWeb web, int klientId, int okresId, string KEY)
        {
            SPList list = web.Lists.TryGetList(targetList);
            SPListItem item = list.Items.Cast<SPListItem>()
                .Where(i => i["KEY"].ToString() == KEY)
                .FirstOrDefault();
            if (item != null)
            {
                return item.ID;
            }
            else
            {
                SPListItem newItem = list.AddItem();
                newItem["KEY"] = KEY;
                newItem["selKlient"] = klientId;
                newItem["selOkres"] = okresId;

                BLL.Models.Klient k = new Models.Klient(web, klientId);

                newItem["enumRozliczeniePD"] = k.RozliczeniePD;
                newItem["enumRozliczenieVAT"] = k.RozliczenieVAT;
                newItem["colFormaOpodatkowaniaPD"] = k.FormaOpodatkowaniaPD;
                newItem["colFormaOpodatkowaniaVAT"] = k.FormaOpodatkowaniaVAT;
                newItem["colFormaOpodakowania_ZUS"] = k.FormaOpodatkowaniaZUS;

                //ustaw CT
                if (k.TypKlienta == "KSH") newItem["ContentType"] = "Karta kontrolna KSH";
                else newItem["ContentType"] = "Karta kontrolna KPiR";

                newItem.SystemUpdate();

                return newItem.ID;
            }
        }

        #region Helpers
        private static string Create_KEY(Microsoft.SharePoint.SPListItem item)
        {
            int klientId = Get_LookupId(item, "selKlient");
            int okresId = Get_LookupId(item, "selOkres");
            return Create_KEY(klientId, okresId);
        }

        public static string Create_KEY(int klientId, int okresId)
        {
            return string.Format(@"{0}::{1}", klientId.ToString(), okresId.ToString());
        }

        private static string Get_LookupValue(Microsoft.SharePoint.SPListItem item, string col)
        {
            return item[col] != null ? new SPFieldLookupValue(item[col].ToString()).LookupValue : string.Empty;
        }

        private static int Get_LookupId(Microsoft.SharePoint.SPListItem item, string col)
        {
            return item[col] != null ? new SPFieldLookupValue(item[col].ToString()).LookupId : 0;
        }

        private static void Copy_Field(SPListItem item, string col0, SPListItem form, string col1)
        {
            form[col1] = item[col0];
        }

        private static void Copy_Field(SPListItem item, SPListItem form, string col)
        {
            form[col] = item[col];
        }

        private static void Copy_Field(SPListItem item, SPListItem form, string col, int mnoznik)
        {
            form[col] = double.Parse(item[col].ToString()) * mnoznik;
        }

        private static void Copy_Id(SPListItem item, SPListItem form, string col)
        {
            form[col] = item.ID;
        }

        #endregion


        public static void Set_PotwierdzenieOdbioruDokumentow(SPWeb web, int klientId, int okresId)
        {
            string KEY = Create_KEY(klientId, okresId);
            int kkId = Get_KartaKontrolnaId(web, klientId, okresId, KEY);
            SPListItem item = Get_KartaKontrolnaById(web, kkId);
            item[colPOTWIERDZENIE_ODBIORU_DOKUMENTOW] = true;
            item.SystemUpdate();

        }

        internal static double Get_WartoscNadwyzkiDoPrzeniesienia(SPWeb web, int klientId, int okresId)
        {
            SPList list = web.Lists.TryGetList(targetList);

            string KEY = BLL.tabKartyKontrolne.Create_KEY(klientId, okresId);
            int kkId = Get_KartaKontrolnaId(web, klientId, okresId, KEY);

            if (kkId>0)
            {
                SPListItem item = Get_KartaKontrolnaById(web, kkId);
                if (item!=null)
                {
                    return BLL.Tools.Get_Value(item, "colVAT_WartoscDoPrzeniesienia");
                }
            }

            return 0;
        }
    }
}
