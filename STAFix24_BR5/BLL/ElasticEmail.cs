using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Collections.Specialized;
using Microsoft.SharePoint;

namespace ElasticEmail
{

    public class EmailGenerator

    {
        //e9b604ee-e197-44a3-a0fa-6b23cf9ec0bb - profil: stafix24@hotmail.com
        //46779a46-20ab-4670-99b8-b44a9a6f45b5 - profil: biuro@rawcom24.pl

        const string USERNAME = "e9b604ee-e197-44a3-a0fa-6b23cf9ec0bb";
        const string API_KEY = "e9b604ee-e197-44a3-a0fa-6b23cf9ec0bb";

        public static string ReportError(Exception ex, string webUrl, string extraInfo)
        {
            string subject = string.Format(":: ERR :: {0}", webUrl);
            StringBuilder sb = new StringBuilder();
            sb.Append(@"<table>");
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "Source", ex.Source);
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "TargetSite", ex.TargetSite);
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "Message", ex.Message);
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "StackTrace", ex.StackTrace);
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "InnerException", ex.InnerException);
            sb.AppendFormat(@"<tr valign='top'><td>{0}</td><td>{1}</td></tr>", "HelpLink", ex.HelpLink);
            sb.Append(@"</table>");

            if (!string.IsNullOrEmpty(extraInfo))
            {
                sb.AppendFormat(@"<div>{0}</div>", extraInfo);
            }

            return ElasticEmail.EmailGenerator.SendMail(subject, sb.ToString());

        }

        public static string ReportError(Exception ex, string webUrl)
        {
            return ReportError(ex, webUrl, string.Empty);
        }
        
        public static string SendMail(string subject, string bodyHtml)
        {

            string from = "mailer@stafix24.pl";
            string fromName = "STAFix24 Mailer";
            string to = "jacek.rawiak@hotmail.com";
            string bodyText = "Text Body";

            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection();
            values.Add("username", USERNAME);
            values.Add("api_key", API_KEY);
            values.Add("from", from);
            values.Add("from_name", fromName);
            values.Add("subject", subject);
            if (bodyHtml != null)
                values.Add("body_html", bodyHtml);
            if (bodyText != null)
                values.Add("body_text", bodyText);
            values.Add("to", to);

            byte[] response = client.UploadValues("https://api.elasticemail.com/mailer/send", values);
            return Encoding.UTF8.GetString(response);
        }

        public static void SendProcessEndConfirmationMail(string subject, string bodyHtml, SPWeb web, SPItem item)
        {
            string from = "robot@stafix24.pl";
            string fromName = "STAFix24 Robot";
            string to = new SPFieldUserValue(web, item["Author"].ToString()).User.Email;
            string bodyText = string.Empty;

            WebClient client = new WebClient();
            NameValueCollection values = new NameValueCollection();
            values.Add("username", USERNAME);
            values.Add("api_key", API_KEY);
            values.Add("from", from);
            values.Add("from_name", fromName);
            values.Add("subject", subject);
            if (bodyHtml != null)
                values.Add("body_html", bodyHtml);
            if (bodyText != null)
                values.Add("body_text", bodyText);
            values.Add("to", to);

            byte[] response = client.UploadValues("https://api.elasticemail.com/mailer/send", values);

        }


    }
}
