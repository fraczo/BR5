using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class KontaKlienta
    {
        private SPWeb web;
        private int klientId;

        const string targetList = "Klienci";

        public KontaKlienta(SPWeb web, int klientId)
        {
            this.web = web;
            this.klientId = klientId;

            SPList list = web.Lists.TryGetList(targetList);
            if (list != null)
            {
                int urzadId = 0;

                SPListItem item = list.GetItemById(klientId);
                if (item != null)
                {

                    if (item["selUrzadSkarbowy"] != null)
                    {
                        try
                        {
                            urzadId = new SPFieldLookupValue(item["selUrzadSkarbowy"].ToString()).LookupId;
                            this.NazwaUrzeduSkarbowego = new SPFieldLookupValue(item["selUrzadSkarbowy"].ToString()).LookupValue;
                            this.NazwaUrzeduSkarbowegoVAT = NazwaUrzeduSkarbowego;
                            this.IdUrzeduSkarbowego = urzadId;
                            this.IdUrzeduSkarbowegoVAT = urzadId;

                            this.KontoPIT = dicUrzedySkarbowe.Get_KontoPIT(web, urzadId);
                            this.KontoCIT = dicUrzedySkarbowe.Get_KontoCIT(web, urzadId);
                            this.KontoVAT = dicUrzedySkarbowe.Get_KontoVAT(web, urzadId);

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

                    //!!! DOTYCZY TYLKO KPiR

                    if (item.ContentType.Name == "KPiR")
                    {
                        if (item["selUrzadSkarbowyVAT"] != null) // czy jest dedykowany urząd skarbowy do rozliczeń VAT
                        {
                            try
                            {
                                urzadId = new SPFieldLookupValue(item["selUrzadSkarbowyVAT"].ToString()).LookupId;
                                this.NazwaUrzeduSkarbowegoVAT = new SPFieldLookupValue(item["selUrzadSkarbowyVAT"].ToString()).LookupValue;

                                string konto = dicUrzedySkarbowe.Get_KontoVAT(web, urzadId);
                                if (!string.IsNullOrEmpty(konto))
                                {
                                    this.KontoVAT = konto;
                                    this.IdUrzeduSkarbowego = urzadId;
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
                }
            }
        }

        public object KontoPIT { get; set; }
        public object KontoCIT { get; set; }
        public string KontoVAT { get; set; }

        public string NazwaUrzeduSkarbowego { get; set; }
        public string NazwaUrzeduSkarbowegoVAT { get; set; }

        public object IdUrzeduSkarbowego { get; set; }
        public object IdUrzeduSkarbowegoVAT { get; set; }


    }
}
