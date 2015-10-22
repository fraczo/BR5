using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Workflow;

namespace TimerJobs
{

    public class ObslugaWiadomosciTJ : Microsoft.SharePoint.Administration.SPJobDefinition
    {

        public static void CreateTimerJob(SPSite site)
        {
            var timerJob = new ObslugaWiadomosciTJ(site);
            timerJob.Schedule = new SPMinuteSchedule
            {
                BeginSecond = 0,
                EndSecond = 0,
                Interval = 15
            };

            timerJob.Update();
        }

        public static void DelteTimerJob(SPSite site)
        {
            site.WebApplication.JobDefinitions
                .OfType<ObslugaWiadomosciTJ>()
                .Where(i => string.Equals(i.SiteUrl, site.Url, StringComparison.InvariantCultureIgnoreCase))
                .ToList()
                .ForEach(i => i.Delete());
        }

        public ObslugaWiadomosciTJ()
            : base()
        {

        }

        public ObslugaWiadomosciTJ(SPSite site)
            : base(string.Format("Biuromagda_Obsługa wiadomości Timer Job ({0})", site.Url), site.WebApplication, null, SPJobLockType.Job)
        {
            Title = Name;
            SiteUrl = site.Url;
        }

        public string SiteUrl
        {
            get { return (string)this.Properties["SiteUrl"]; }
            set { this.Properties["SiteUrl"] = value; }
        }

        public override void Execute(Guid targetInstanceId)
        {
            using (var site = new SPSite(SiteUrl))
            {
                SPList list = site.RootWeb.Lists.TryGetList("admProcessRequests");

                SPListItem item = list.AddItem();
                item["ContentType"] = "Obsługa wiadomości";
                item.SystemUpdate();
            }
        }

    }
}


