using System;
using System.Security.Permissions;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.Workflow;
using BLL;
using ElasticEmail;
using EventReceivers.admProcessRequestsER;

namespace admProcessRequests_EventReceiver
{
    /// <summary>
    /// List Item Events
    /// </summary>
    public class EventReceiver1 : SPItemEventReceiver
    {
        /// <summary>
        /// An item was added.
        /// </summary>
        public override void ItemAdded(SPItemEventProperties properties)
        {
            Execute(properties);
        }

        /// <summary>
        /// An item was updated.
        /// </summary>
        public override void ItemUpdated(SPItemEventProperties properties)
        {
            Execute(properties);
        }


        private void Execute(SPItemEventProperties properties)
        {
            this.EventFiringEnabled = false;

            properties.ListItem["enumStatusZlecenia"] = "Obsługa";
            properties.ListItem.SystemUpdate();

            try
            {
                SPSecurity.RunWithElevatedPrivileges(delegate()
                {
                    using (SPSite site = new SPSite(properties.SiteId))
                    {
                        using (SPWeb web = site.AllWebs[properties.Web.ID])
                        {
                            //określ rodzaj raportu
                            string ct = properties.ListItem["ContentType"].ToString();
                            switch (ct)
                            {
                                case "Generowanie formatek rozliczeniowych":
                                    GeneratorFormatekRozliczeniowych.Execute_GenFormRozl(properties, web);
                                    //PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    //obsługa wewnętrz porcedury
                                    break;
                                case "Generowanie formatek rozliczeniowych dla klienta":
                                    GeneratorFormatekRozliczeniowych.Execute_GenFormRozlK(properties, web);
                                    //PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    //obsługa wewnętrz porcedury
                                    break;
                                case "Import Klientów":
                                    string message;
                                    ImportKlientow.Execute(properties, web, out message);
                                    PotwierdzMailemZakonczenieZlecenia(properties, web, ct, message);
                                    break;
                                case "Import bufora wiadomości":
                                    ImportBuforaWiadomosci.Execute(properties, web);
                                    PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    break;
                                case "Import faktur elektronicznych":
                                    ImportFakturElektronicznych.Execute(properties, web);
                                    PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    break;
                                case "Import faktur za obsługę":
                                    ImportFakturZaObsluge.Execute(properties, web);
                                    PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    break;
                                case "Import przeterminowanych należności":
                                    ImportPrzeterminowanychNaleznosci.Execute(properties, web);
                                    PotwierdzMailemZakonczenieZlecenia(properties, web, ct);
                                    break;
                                case "Obsługa wiadomości":
                                    ObslugaWiadomosci.Execute(properties, web);
                                    break;
                                case "Obsługa zadań":
                                    ObslugaZadan.Execute(properties, web);
                                    break;
                                case "Obsługa ADO":
                                    ObslugaADO.Execute(properties, web);
                                    break;
                                default:
                                    //properties.ListItem["colStatus"] = "Zakończony";
                                    //properties.ListItem.SystemUpdate();
                                    break;
                            }

                        }
                    }

                });
            }
            catch (Exception ex)
            {
                properties.ListItem["enumStatusZlecenia"] = "Anulowany";
                properties.ListItem.SystemUpdate();

                BLL.Logger.LogEvent(properties.WebUrl, ex.ToString());
                var result = ElasticEmail.EmailGenerator.ReportError(ex, properties.WebUrl.ToString());

            }
            finally
            {
                properties.ListItem.Delete();
//                properties.ListItem["enumStatusZlecenia"] = "Zakończony";
//                properties.ListItem.SystemUpdate();

//                //oczyszczanie listy zleceń starszych niż 1 dzień
//                try
//                {
//                    BLL.admProcessRequests.List_Cleanup(properties.Web, 1);
//                }
//                catch (Exception)
//                {
//#if DEBUG
//                    throw;
//#endif
//                }              

                this.EventFiringEnabled = true;
            }
        }



        private static void PotwierdzMailemZakonczenieZlecenia(SPItemEventProperties properties, SPWeb web, string ct)
        {
            string bodyHtml = String.Format("zlecono {0}", properties.ListItem["Created"].ToString());
            PotwierdzMailemZakonczenieZlecenia(properties, web, ct, bodyHtml);
        }

        private static void PotwierdzMailemZakonczenieZlecenia(SPItemEventProperties properties, SPWeb web, string ct, string bodyHtml)
        {
            string subject = ct.ToString();
#if DEBUG
            //send directly via ElasticEmail
            ElasticEmail.EmailGenerator.SendProcessEndConfirmationMail(
                subject,
                bodyHtml,
                web,
                properties.ListItem);

#else
                                    //send via SPUtility
                                    SPEmail.EmailGenerator.SendProcessEndConfirmationMail(
                                        subject,
                                        bodyHtml,
                                        web,
                                        properties.ListItem);
#endif
        }
    }
}
