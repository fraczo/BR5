using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class BiuroRachunkowe
    {
        const string targetList = @"admSetup";
        private SPWeb web;
        private int okresId;

        const string listaOkresow = "Okresy";

        public BiuroRachunkowe(SPWeb web)
        {
            this.web = web;

            Get_BRData();

        }
        public BiuroRachunkowe(SPWeb web, int okresId)
        {
            this.web = web;
            this.okresId = okresId;

            Get_BRData();
            Get_TerminPlatnosci();
        }

        #region Helpers

        private void Get_BRData()
        {
            SPList list = this.web.Lists.TryGetList(targetList);
            if (list != null)
            {
                list.Items.Cast<SPListItem>()
                    .ToList()
                    .ForEach(item =>
                    {
                        string key = item["KEY"].ToString();
                        switch (key)
                        {
                            case @"BR_MIEJSCOWOSC":
                                Miejscowosc = item["VALUE"].ToString();
                                break;
                            case @"BR_KOD_POCZTOWY":
                                KodPocztowy = item["VALUE"].ToString();
                                break;
                            case @"BR_ADRES":
                                Adres = item["VALUE"].ToString();
                                break;
                            case @"BR_NAZWA":
                                Nazwa = item["VALUE"].ToString();
                                break;
                            case @"BR_KONTO":
                                Konto = item["VALUE"].ToString();
                                break;

                            default:
                                break;
                        }
                    });
            }
        }
        private void Get_TerminPlatnosci()
        {
            SPList list = this.web.Lists.TryGetList(listaOkresow);
            if (list != null)
            {
                SPListItem item = list.GetItemById(this.okresId);
                if (item != null)
                {
                    if (item["colBR_TerminPrzekazania"]!=null)
                    {
                        TerminPrzekazania = DateTime.Parse(item["colBR_TerminPrzekazania"].ToString());
                    }
                }
            }
        }

        #endregion

        public string Nazwa { get; set; }
        public string  Adres { get; set; }
        public string  KodPocztowy { get; set; }
        public string  Miejscowosc { get; set; }
        public string Konto { get; set; }
        public DateTime TerminPrzekazania { get; set; }
    }
}
