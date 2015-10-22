using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class UrzadSkarbowy
    {
        public string NazwaUrzedu { get; set; }
        public string Adres { get; set; }
        public string KodPocztowy { get; set; }
        public string Miejscowosc { get; set; }

        const string targetList = "Urzędy skarbowe"; //"dicUrzedySkarbowe";

        public UrzadSkarbowy(SPWeb web, int urzadId)
        {
            SPList list = web.Lists.TryGetList(targetList);

            SPListItem item = list.GetItemById(urzadId);
            if (item != null)
            {
                NazwaUrzedu = item.Title;
                Adres = item["colAdres"] != null ? item["colAdres"].ToString() : string.Empty;
                KodPocztowy = item["colKodPocztowy"] != null ? item["colKodPocztowy"].ToString() : string.Empty;
                Miejscowosc = item["colMiejscowosc"] != null ? item["colMiejscowosc"].ToString() : string.Empty;
            }

        }

        public string Get_NazwaOdbiorcyPrzelewu()
        {
            string s = NazwaUrzedu + " " + Adres + " " + KodPocztowy + " " + Miejscowosc;
            s = s.Replace("  ", " ").Replace("  ", " ").Replace("  ", " ").Trim();
            return s;
        }
    }
}
