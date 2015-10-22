using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;

namespace BLL.Models
{
    public class Flagi
    {
        public Flagi(SPWeb web, int klientId)
        {

            PrzypomnienieOTerminiePlatnosci = false;
            GenerowanieDrukuWplaty = false;
            AudytDanych = false;


            if (tabKlienci.HasServiceAssigned(web, klientId, "POT"))
            {
                PrzypomnienieOTerminiePlatnosci = true;
            }

            if (tabKlienci.HasServiceAssigned(web, klientId, "GBW"))
            {
                GenerowanieDrukuWplaty = true;
            }

            if (tabKlienci.HasServiceAssigned(web, klientId, "AD") || tabKlienci.HasServiceAssigned(web, klientId, "ADO"))
            {
                AudytDanych = true;
            }
            
        }

        public bool PrzypomnienieOTerminiePlatnosci { get; set; }
        public bool GenerowanieDrukuWplaty { get; set; }
        public bool AudytDanych { get; set; }



    }
}
