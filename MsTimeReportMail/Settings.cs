using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MsTimeReportMail
{
    public partial class Settings : Form
    {
        public Settings()
        {
            InitializeComponent();

            //load settings: ToRecipient, CcRecipient, Body
            if (Properties.Settings.Default != null)
            {
                TbxTo.Text = Properties.Settings.Default["ToRecipient"].ToString();
                TbxCc.Text = Properties.Settings.Default["CcRecipient"].ToString();
                TbxBody.Text = Properties.Settings.Default["Body"].ToString();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //save settings
            Properties.Settings.Default["ToRecipient"] = TbxTo.Text;
            Properties.Settings.Default["CcRecipient"] = TbxCc.Text;
            Properties.Settings.Default["Body"] = TbxBody.Text;

            Properties.Settings.Default.Save();
            this.Close();
        }

       

        
    }
}
