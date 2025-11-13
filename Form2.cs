using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DAT_ToolReports
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            cbxHour1.Text  = DAT_ToolReports.Properties.Settings.Default.NightHour1.ToString();
            cbxHour2.Text = DAT_ToolReports.Properties.Settings.Default.NightHour2.ToString();
            cbxMinute1.Text = DAT_ToolReports.Properties.Settings.Default.NightMinute1.ToString();
            cbxMinute2.Text = DAT_ToolReports.Properties.Settings.Default.NightMinute2.ToString();
            txtCentre.Text = DAT_ToolReports.Properties.Settings.Default.Centre;
            txtCompany.Text = DAT_ToolReports.Properties.Settings.Default.Company;
            txtLinkSever.Text = DAT_ToolReports.Properties.Settings.Default.linkServer;
            txtProvince.Text = DAT_ToolReports.Properties.Settings.Default.Province;
            ckbCountATforNight.Checked = DAT_ToolReports.Properties.Settings.Default.CountATforNight;
            ckbNightByStart.Checked = DAT_ToolReports.Properties.Settings.Default.NightByStart;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            DAT_ToolReports.Properties.Settings.Default.Centre = txtCentre.Text;
            DAT_ToolReports.Properties.Settings.Default.Company = txtCompany.Text;
            DAT_ToolReports.Properties.Settings.Default.linkServer = txtLinkSever.Text;
            DAT_ToolReports.Properties.Settings.Default.Province = txtProvince.Text;
            DAT_ToolReports.Properties.Settings.Default.NightHour1 = Int32.Parse(cbxHour1.Text);
            DAT_ToolReports.Properties.Settings.Default.NightHour2 = Int32.Parse(cbxHour2.Text);
            DAT_ToolReports.Properties.Settings.Default.NightMinute1 = Int32.Parse(cbxMinute1.Text);
            DAT_ToolReports.Properties.Settings.Default.NightMinute2 = Int32.Parse(cbxMinute2.Text);
            DAT_ToolReports.Properties.Settings.Default.CountATforNight = ckbCountATforNight.Checked;
            DAT_ToolReports.Properties.Settings.Default.NightByStart = ckbNightByStart.Checked;
            DAT_ToolReports.Properties.Settings.Default.Save();
            this.DialogResult = DialogResult.OK;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
