using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class Klient
    {
        private Microsoft.SharePoint.SPWeb web;
        private int klientId;

        public Klient(Microsoft.SharePoint.SPWeb web, int klientId)
        {

            this.web = web;
            this.klientId = klientId;

            SPListItem item = tabKlienci.Get_KlientById(web, klientId);

            if (item != null)
            {
                this.TypKlienta = item.ContentType.Name;
                OddzialZUSId = item["selOddzialZUS"] != null ? new SPFieldLookupValue(item["selOddzialZUS"].ToString()).LookupId : 0;
                FormaOpodatkowaniaZUS = item["colFormaOpodakowania_ZUS"] != null ? item["colFormaOpodakowania_ZUS"].ToString() : string.Empty;
                FormaOpodatkowaniaKPiR = item["colFormaOpodatkowaniaPD_KPiR"] != null ? item["colFormaOpodatkowaniaPD_KPiR"].ToString() : string.Empty;
                FormaOpodatkowaniaKSH = item["colFormaOpodatkowaniaPD_KSH"] != null ? item["colFormaOpodatkowaniaPD_KSH"].ToString() : string.Empty;
                FormaOpodatkowaniaVAT = item["colFormaOpodatkowaniaVAT"] != null ? item["colFormaOpodatkowaniaVAT"].ToString() : string.Empty;

                RozliczeniePD = item["enumRozliczeniePD"] != null ? item["enumRozliczeniePD"].ToString() : string.Empty;
                RozliczenieVAT = item["enumRozliczenieVAT"] != null ? item["enumRozliczenieVAT"].ToString() : string.Empty;


                string ct = item["ContentType"].ToString();
                //Forma opodatkowania
                switch (ct)
                {
                    case "KPiR":
                    case "Osoba fizyczna":
                        FormaOpodatkowaniaPD = FormaOpodatkowaniaKPiR;
                        break;
                    case "KSH":
                    case "Firma":
                        FormaOpodatkowaniaPD = FormaOpodatkowaniaKSH;
                        break;
                    default:
                        FormaOpodatkowaniaPD = string.Empty;
                        break;
                }

                //Pełna nazwa firmy
                switch (ct)
                {
                    case "KPiR":
                        this.FormaPrawna = item["enumFormaPrawna_KPiR"] != null ? item["enumFormaPrawna_KPiR"].ToString() : string.Empty;
                        this.PelnaNazwaFirmy = item.Title;
                        break;
                    case "KSH":
                        this.PelnaNazwaFirmy = item.Title;
                        this.FormaPrawna = item["enumFormaPrawna"] != null ? item["enumFormaPrawna"].ToString() : string.Empty;
                        break;
                    case "Osoba fizyczna":

                        this.PelnaNazwaFirmy = string.Format("{0} {1} {2}",
                            item["colImie"] != null ? item["colImie"].ToString() : string.Empty,
                            item["colNazwisko"] != null ? item["colNazwisko"].ToString() : string.Empty,
                            item["colPESEL"] != null ? "pesel: " + item["colPESEL"].ToString() : string.Empty).Trim();
                        break;
                    case "Firma":
                        this.PelnaNazwaFirmy = item["colNazwa"] != null ? item["colNazwa"].ToString() : string.Empty;
                        break;
                    default:
                        this.PelnaNazwaFirmy = string.Empty;
                        break;
                }

                OsobaDoKontaktu = item["colOsobaDoKontaktu"] != null ? item["colOsobaDoKontaktu"].ToString() : string.Empty;
                Email = item["colEmail"] != null ? item["colEmail"].ToString() : string.Empty;
                Telefon = item["colTelefon"] != null ? item["colTelefon"].ToString() : string.Empty;
                Adres = item["colAdres"] != null ? item["colAdres"].ToString() : string.Empty;
                KodPocztowy = item["colKodPocztowy"] != null ? item["colKodPocztowy"].ToString() : string.Empty;
                Miejscowosc = item["colMiejscowosc"] != null ? item["colMiejscowosc"].ToString() : string.Empty;
                NIP = item["colNIP"] != null ? item["colNIP"].ToString() : string.Empty;
                NazwaFirmy = item.Title;
                Regon = item["colRegon"] != null ? item["colRegon"].ToString() : string.Empty;

                this.UwagiKadrowe = item["colUwagiKadrowe"] != null ? item["colUwagiKadrowe"].ToString() : string.Empty;
                this.Uwagi = item["colUwagi"] != null ? item["colUwagi"].ToString() : string.Empty;

                this.ZatrudniaPracownikow = item["colZatrudniaPracownikow"] != null ? bool.Parse(item["colZatrudniaPracownikow"].ToString()) : false ;


                if (item["selTerminPlatnosci"] != null)
                {
                    int terminPlatnosciId = new SPFieldLookupValue(item["selTerminPlatnosci"].ToString()).LookupId;
                    this.TerminPlatnosci = dicTerminyPlatnosci.Get_TerminPlatnosci(web, terminPlatnosciId);
                }
                else
                {
                    this.TerminPlatnosci = 0;
                }

                //operatorzy
                this.OperatorId_Audyt = item["selDedykowanyOperator_Audyt"] != null ? new SPFieldLookupValue(item["selDedykowanyOperator_Audyt"].ToString()).LookupId : 0;
                this.OperatorId_Podatki = item["selDedykowanyOperator_Podatki"] != null ? new SPFieldLookupValue(item["selDedykowanyOperator_Podatki"].ToString()).LookupId : 0;
                this.OperatorId_Kadry = item["selDedykowanyOperator_Kadry"] != null ? new SPFieldLookupValue(item["selDedykowanyOperator_Kadry"].ToString()).LookupId : 0;

                //Daty
                this.DataRozpoczeciaDzialalnosci = BLL.Tools.Get_Date(item,"colDataRozpoczeciaDzialalnosci");


                // PIT
                try
                {
                    int urzadId = item["selUrzadSkarbowy"] != null ? new SPFieldLookupValue(item["selUrzadSkarbowy"].ToString()).LookupId : 0;
                    urzadId = BLL.dicUrzedySkarbowe.Ensure(web, urzadId);
                    if (urzadId > 0)
                    {
                        if (this.FormaOpodatkowaniaPD=="CIT")
                            NumerRachunkuPD = tabUrzedySkarbowe.Get_NumerRachunkuCITById(web, urzadId);
                        else
                            NumerRachunkuPD = tabUrzedySkarbowe.Get_NumerRachunkuPITById(web, urzadId);

                        NazwaUrzeduSkarbowego = tabUrzedySkarbowe.Get_NazwaUrzeduById(web, urzadId);
                        UrzadSkarbowyId = urzadId;
                    }
                    else
                    {
                        NumerRachunkuPD = string.Empty;
                        UrzadSkarbowyId = 0;
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                throw ex;
#else
                    BLL.Logger.LogEvent(web.Url, ex.ToString() + " KlientId= " + klientId.ToString());
                    var result = ElasticEmail.EmailGenerator.ReportError(ex, web.Url, BLL.Tools.Get_ItemInfo(item));
#endif

                }

                //VAT

                try
                {
                    int urzadVATId = item["selUrzadSkarbowyVAT"] != null ? new SPFieldLookupValue(item["selUrzadSkarbowyVAT"].ToString()).LookupId : 0;
                    urzadVATId = BLL.dicUrzedySkarbowe.Ensure(web, urzadVATId);
                    if (urzadVATId > 0)
                    {
                        NumerRachunkuVAT = tabUrzedySkarbowe.Get_NumerRachunkuVATById(web, urzadVATId);
                        NazwaUrzeduSkarbowegoVAT = tabUrzedySkarbowe.Get_NazwaUrzeduById(web, urzadVATId);
                        UrzadSkarbowyVATId = urzadVATId;
                    }
                    else
                    {
                        //Przyjmij parametry jak dla US od podatku PIT
                        NumerRachunkuVAT = this.NumerRachunkuPD;
                        NazwaUrzeduSkarbowegoVAT = this.NazwaUrzeduSkarbowego;
                        UrzadSkarbowyVATId = this.UrzadSkarbowyId;
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                throw ex;
#else
                    BLL.Logger.LogEvent(web.Url, ex.ToString() + " KlientId= " + klientId.ToString());
                    var result = ElasticEmail.EmailGenerator.ReportError(ex, web.Url, BLL.Tools.Get_ItemInfo(item));
#endif
                }

                // ZUS

                try
                {
                    int oddzialZUSId = item["selOddzialZUS"] != null ? new SPFieldLookupValue(item["selOddzialZUS"].ToString()).LookupId : 0;
                    oddzialZUSId = BLL.dicOddzialyZUS.Ensure(web, oddzialZUSId);
                    if (oddzialZUSId > 0)
                    {
                        this.OddzialZUSId = oddzialZUSId;
                    }
                    else
                    {
                        //Przyjmij parametry jak dla US od podatku PIT
                        NumerRachunkuVAT = this.NumerRachunkuPD;
                        NazwaUrzeduSkarbowegoVAT = this.NazwaUrzeduSkarbowego;
                        UrzadSkarbowyVATId = this.UrzadSkarbowyId;
                    }
                }
                catch (Exception ex)
                {
#if DEBUG
                    throw ex;
#else
                    BLL.Logger.LogEvent(web.Url, ex.ToString() + " KlientId= " + klientId.ToString());
                    var result = ElasticEmail.EmailGenerator.ReportError(ex, web.Url, BLL.Tools.Get_ItemInfo(item));
#endif
                }

            }

        }

        public string FormaOpodatkowaniaZUS { get; set; }
        public string FormaOpodatkowaniaKPiR { get; set; }
        public string FormaOpodatkowaniaKSH { get; set; }

        public string FormaOpodatkowaniaPD { get; set; }

        public string FormaOpodatkowaniaVAT { get; set; }
        public string OsobaDoKontaktu { get; set; }
        public string Email { get; set; }
        public string Telefon { get; set; }
        public string Adres { get; set; }
        public string KodPocztowy { get; set; }
        public string Miejscowosc { get; set; }
        public int OddzialZUSId { get; set; }
        public int TerminPlatnosci { get; set; }
        public string NIP { get; set; }

        public string NazwaFirmy { get; set; }

        public string Regon { get; set; }

        public string NumerRachunkuPD { get; set; }

        public string NazwaUrzeduSkarbowego { get; set; }

        public string Get_NazwaNadawcyPrzelewu()
        {
            string s = NazwaFirmy + " " + Adres + " " + KodPocztowy + " " + Miejscowosc;
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim();
            return s;

        }

        public int UrzadSkarbowyId { get; set; }
        public string UwagiKadrowe { get; set; }

        public string RozliczeniePD { get; set; }
        public string RozliczenieVAT { get; set; }

        public int OperatorId_Audyt { get; set; }
        public int OperatorId_Podatki { get; set; }
        public int OperatorId_Kadry { get; set; }

        public string PelnaNazwaFirmy { get; set; }

        public string TypKlienta { get; set; }

        public string Uwagi { get; set; }

        public DateTime DataRozpoczeciaDzialalnosci { get; set; }

        public string FormaPrawna { get; set; }

        public bool ZatrudniaPracownikow { get; set; }

        public string NumerRachunkuVAT { get; set; }

        public int UrzadSkarbowyVATId { get; set; }

        public string NazwaUrzeduSkarbowegoVAT { get; set; }
    }
}
