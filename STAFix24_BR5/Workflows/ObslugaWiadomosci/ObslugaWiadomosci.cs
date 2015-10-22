using System;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Workflow.ComponentModel.Compiler;
using System.Workflow.ComponentModel.Serialization;
using System.Workflow.ComponentModel;
using System.Workflow.ComponentModel.Design;
using System.Workflow.Runtime;
using System.Workflow.Activities;
using System.Workflow.Activities.Rules;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Workflow;
using Microsoft.SharePoint.WorkflowActions;
using System.Net.Mail;
using System.Net;
using System.Collections.Specialized;
using System.Text;

namespace Workflows.ObslugaWiadomosci
{
    public sealed partial class ObslugaWiadomosci : SequentialWorkflowActivity
    {
        public ObslugaWiadomosci()
        {
            InitializeComponent();
        }

        public Guid workflowId = default(System.Guid);
        public SPWorkflowActivationProperties workflowProperties = new SPWorkflowActivationProperties();
        public SPListItem item;
        public MailMessage mail;
        private bool isMailReadyToSend;

        private void onWorkflowActivated1_Invoked(object sender, ExternalDataEventArgs e)
        {
            item = workflowProperties.Item;
        }

        private void Mail_Setup_ExecuteCode(object sender, EventArgs e)
        {
            mail = new MailMessage();
            isMailReadyToSend = false;

            //From
            if (item["colNadawca"] != null)
            {
                mail.From = new MailAddress(item["colNadawca"].ToString());
            }
            else
            {
                mail.From = new MailAddress(BLL.admSetup.GetValue(item.ParentList.ParentWeb, "EMAIL_BIURA"));
            }

            //To
            if (item["colOdbiorca"] != null && !string.IsNullOrEmpty(item.Title))
            {
                mail.To.Add(new MailAddress(item["colOdbiorca"].ToString()));

                //CC
                bool isKopiaDoNadawcy = item["colKopiaDoNadawcy"] != null ? (bool)item["colKopiaDoNadawcy"] : false;
                if (isKopiaDoNadawcy)
                {
                    mail.CC.Add(new MailAddress(item["colNadawca"].ToString()));
                }

                //BCC
                bool isKopiaDoBiura = item["colKopiaDoBiura"] != null ? (bool)item["colKopiaDoBiura"] : false;
                if (isKopiaDoBiura)
                {
                    mail.Bcc.Add(new MailAddress(BLL.admSetup.GetValue(item.ParentList.ParentWeb, "EMAIL_BIURA_ARCH")));
                }

                //Subject
                mail.Subject = item.Title;

                //Body
                if (item["colTrescHTML"] != null)
                {
                    string bodyHTML = item["colTrescHTML"].ToString();
                    mail.Body = bodyHTML;
                    mail.IsBodyHtml = true;
                }
                else
                {
                    if (item["colTresc"] != null)
                    {
                        string body = item["colTresc"].ToString();

                        StringBuilder sb = new StringBuilder(BLL.dicSzablonyKomunikacji.Get_TemplateByKod(item,"EMAIL_DEFAULT_BODY",true));
                        sb.Replace(@"___BODY___", body);
                        sb.Replace(@"___FOOTER___", string.Empty);
                        mail.Body = sb.ToString();
                        mail.IsBodyHtml = true;
                    }
                }

                isMailReadyToSend = true;
            }
        }

        private void Mail_Send_ExecuteCode(object sender, EventArgs e)
        {
            if (isMailReadyToSend)
            {
                bool testMode = true;
                SPEmail.EmailGenerator.SendMailFromMessageQueue(item, mail, testMode);

                //ustaw flagę wysyłki
                item["colCzyWyslana"] = true;
                item["colDataNadania"] = DateTime.Now.ToString();
                item.SystemUpdate();

                int zadanieId = item["_ZadanieId"] != null ? int.Parse(item["_ZadanieId"].ToString()) : 0;
                if (zadanieId > 0)
                {
                    BLL.tabZadania.Update_StatusWysylki(item.Web, item, zadanieId, BLL.Models.StatusZadania.Zakończone);
                }
            }
        }

        private void isWiadomoscWyslana(object sender, ConditionalEventArgs e)
        {
            e.Result = item["czyWyslana"] != null ? bool.Parse(item["czyWyslana"].ToString()) : false;
        }

        private void Update_tabKartyKontrolne_ExecuteCode(object sender, EventArgs e)
        {
            int zadanieId = item["_ZadanieId"] != null ? int.Parse(item["_ZadanieId"].ToString()) : 0;
            if (zadanieId > 0)
            {
                SPListItem task = BLL.tabZadania.Get_ZadanieById(item.Web, zadanieId);
                if (task != null)
                {
                    DateTime date = DateTime.Parse(item["Modified"].ToString());
                    string ct = task.ContentType.Name;
                    switch (ct)
                    {
                        case "Rozliczenie z biurem rachunkowym":
                            BLL.tabZadania.Update_RBR_DataWysylki(task, date);
                            BLL.tabKartyKontrolne.Update_RBR_DataWysylki(task, date);
                            break;
                        case "Rozliczenie podatku dochodowego":
                        case "Rozliczenie podatku dochodowego spółki":
                            BLL.tabZadania.Update_PD_DataWysylki(task, date);
                            BLL.tabKartyKontrolne.Update_PD_DataWysylki(task, date);
                            break;
                        case "Rozliczenie podatku VAT":
                            BLL.tabZadania.Update_VAT_DataWysylki(task, date);
                            BLL.tabKartyKontrolne.Update_VAT_DataWysylki(task, date);
                            break;
                        case "Rozliczenie ZUS":
                            BLL.tabZadania.Update_ZUS_DataWysylki(task, date);
                            BLL.tabKartyKontrolne.Update_ZUS_DataWysylki(task, date);
                            break;
                        default:
                            break;
                    }
                }
            }

        }

        private void isFlagaWysłanoUstawiona(object sender, ConditionalEventArgs e)
        {
            e.Result = item["colCzyWyslana"] != null ? (bool)item["colCzyWyslana"] : false;
        }

        private void setState_Anulowana_MethodInvoking(object sender, EventArgs e)
        {
            //SetStatusWysylki(enumStatusWysylki.Anulowana);
        }

        private void setState_Wysłana_MethodInvoking(object sender, EventArgs e)
        {
            SetStatusWysylki(enumStatusWysylki.Wysłana);
        }

        private void setState_PrzygotowanieWysyłki_MethodInvoking(object sender, EventArgs e)
        {
            SetStatusWysylki(enumStatusWysylki.Obsługa);
        }

        private void SetStatusWysylki(enumStatusWysylki status)
        {
            item["enumStatusWysylki"] = status;
            item["colDataNadania"] = DateTime.Now;
            item.SystemUpdate();
        }






    }
}
