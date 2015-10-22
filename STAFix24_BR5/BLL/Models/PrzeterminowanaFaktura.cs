using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class PrzeterminowanaFaktura
    {
        public PrzeterminowanaFaktura(SPListItem record)
        {
            if (record!=null)
            {
                if (record["Title"]!=null)
                {
                    NazwaDluznika = record["Title"].ToString();
                }
                if (record["colNumerFaktury"] != null)
                {
                    NumerFaktury = record["colNumerFaktury"].ToString();
                }
                if (record["Title"] != null)
                {
                    NazwaDluznika = record["Title"].ToString();
                }
                if (record["Title"] != null)
                {
                    NazwaDluznika = record["Title"].ToString();
                }
                if (record["Title"] != null)
                {
                    NazwaDluznika = record["Title"].ToString();
                }
                if (record["Title"] != null)
                {
                    NazwaDluznika = record["Title"].ToString();
                }
            }

        }

        public String NazwaDluznika { get; set; }
        public String NumerFaktury { get; set; }
        public DateTime DataSprzedazy { get; set; }
        public DateTime DataWystawienia { get; set; }
        public DateTime TerminPlatnosci { get; set; }


    }
}
