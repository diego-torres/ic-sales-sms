﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using CommonAdminPaq;

namespace VentasSMS
{
    public partial class frmMain : Form
    {
        private AdminPaqImpl api;
        public AdminPaqImpl API { get { return api; } }

        public frmMain()
        {
            InitializeComponent();
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            api = new AdminPaqImpl();
            api.InitializeSDK();
            this.WindowState = FormWindowState.Minimized;
            this.Hide();
        }

        private void openConfigForm(string fileName) 
        {
            frmConfig fConfig = new frmConfig();
            fConfig.API = api;
            fConfig.textBoxFileName.Text = fileName;
            fConfig.Show();
        }

        private void telefonosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // EVALUATE IF CONFIGURED FILE EXISTS
            // For read access you do not need to call OpenExeConfiguraton
            string configFilePath = ConfigurationManager.AppSettings[CommonConstants.SMS_FILE_NAME];
            string myDocs = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            configFilePath = configFilePath.Replace("~", myDocs);

            if ("".Equals(configFilePath.Trim()) || !File.Exists(configFilePath))
            {
                openConfigForm(configFilePath);
            }
            else 
            {
                System.Diagnostics.Process.Start(configFilePath);
            }

        }
    }
}
