using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using dialog =  System.Windows.Forms;
using System.Globalization;
using System.Windows.Forms;

namespace MsTimeReportMail
{
    public partial class NewTimeReport
    {
        private void NewTimeReport_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Outlook.Application application = Globals.ThisAddIn.Application;
            Outlook.ExchangeUser currentUser = application.Session.CurrentUser.AddressEntry.GetExchangeUser();
            Outlook.MailItem myMailItem = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);

            //attaching files to mail
            dialog.OpenFileDialog attachment = new dialog.OpenFileDialog();
            attachment.Title = "Select files to send";
            attachment.Multiselect = true;
            DialogResult res =  attachment.ShowDialog();

            if (res == DialogResult.OK)
            {
                if (attachment.FileNames.Length > 0)
                {
                    foreach (string file in attachment.FileNames)
                    {
                        myMailItem.Attachments.Add(file, Outlook.OlAttachmentType.olByValue, 1, file);
                    }
                }
                //setting mail properties
                string subjectFormat = "{0} - {1}a {2} {3}";
                myMailItem.To = Properties.Settings.Default["ToRecipient"].ToString();
                myMailItem.CC = Properties.Settings.Default["CcRecipient"].ToString();
                myMailItem.Subject = string.Format(subjectFormat, currentUser.Name, GetWeekNumberinCurrentMonth(), DateTime.Now.ToString("MMMM", CultureInfo.GetCultureInfo("it-IT")), DateTime.Now.Year);
                myMailItem.BodyFormat = OlBodyFormat.olFormatHTML;
                myMailItem.HTMLBody = Properties.Settings.Default["Body"].ToString();
                //open mail message
                myMailItem.Display(true);
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //opens settings dialog form
            Settings st = new Settings();
            st.ShowDialog();
        }


        private int GetWeekNumberinCurrentMonth()
        {

            Calendar gc = CultureInfo.CurrentCulture.Calendar;
            int yearweek  = gc.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstDay, DayOfWeek.Monday);
            int firstmonthweek = gc.GetWeekOfYear(new DateTime(DateTime.Now.Year,DateTime.Now.Month,1), CalendarWeekRule.FirstDay, DayOfWeek.Monday);

            return yearweek - firstmonthweek + 1;
        }
    }
}
