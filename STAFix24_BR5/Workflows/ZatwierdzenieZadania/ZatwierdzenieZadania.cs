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

namespace Workflows.ZatwierdzenieZadania
{
    public sealed partial class ZatwierdzenieZadania : SequentialWorkflowActivity
    {
        public ZatwierdzenieZadania()
        {
            InitializeComponent();
        }

        public Guid workflowId = default(System.Guid);
        public SPWorkflowActivationProperties workflowProperties = new SPWorkflowActivationProperties();

        private void codeExecute_ExecuteCode(object sender, EventArgs e)
        {
            //SPListItem item = workflowProperties.Item;
            SPListItem item = workflowProperties.List.GetItemById(workflowProperties.ItemId); //Może to pomoże wymusić obsługę zdarzeń

            string status = BLL.Tools.Get_Text(item, "enumStatusZadania");
            switch (status)
            {
                case "Nowe":
                case "Obsługa":
                    if (item.ContentType.Name == "Prośba o dokumenty"
                        || item.ContentType.Name == "Prośba o przesłanie wyciągu bankowego"
                        || item.ContentType.Name == "Rozliczenie z biurem rachunkowym")
                        Zatwierdz_Zadanie(item);
                    break;
                case "Gotowe":
                    if (item.ContentType.Name == "Rozliczenie ZUS"
                        || item.ContentType.Name == "Rozliczenie podatku VAT"
                        || item.ContentType.Name == "Rozliczenie podatku dochodowego"
                        || item.ContentType.Name == "Rozliczenie podatku dochodowego spółki")
                        Zatwierdz_Zadanie(item);
                    break;
                default:
                    break;
            }

        }

        private static void Zatwierdz_Zadanie(SPListItem item)
        {
            string cmd = BLL.Tools.Get_Text(item, "cmdFormatka");
            if (string.IsNullOrEmpty(cmd))
            {
                item["cmdFormatka"] = "Zatwierdź";
                item.SystemUpdate();
            }

            EventReceivers.tabZadaniaER.tabZadaniaER o = new EventReceivers.tabZadaniaER.tabZadaniaER();
            o.Execute(item);

        }

        private void codeTriggerItemEventReceiver_ExecuteCode(object sender, EventArgs e)
        {
            try
            {
                //EventReceivers.tabZadaniaER.tabZadaniaER o = new EventReceivers.tabZadaniaER.tabZadaniaER();
                //o.Execute(workflowProperties.Item);
            }
            catch (Exception)
            {
#if DEBUG
                throw;
#endif
            }

        }
    }
}
