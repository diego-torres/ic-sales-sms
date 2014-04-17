using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using CommonAdminPaq.dto;
using CommonAdminPaq;
using VentasSMS.Properties; 

namespace VentasSMS
{
    public partial class frmConfig : Form
    {
        private AdminPaqImpl api;
        public AdminPaqImpl API { get { return api; } set { api = value; } }

        public frmConfig()
        {
            InitializeComponent();
        }

        private void buttonExplore_Click(object sender, EventArgs e)
        {
            openFileDialogPhones.Filter = "Archivos de Excel|*.xlsx;*.xls";
            openFileDialogPhones.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialogPhones.ShowDialog();
            textBoxFileName.Text = openFileDialogPhones.FileName;
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        

        private void saveSettings()
        {
            Settings set = Settings.Default;
            toolStripProgressBar1.PerformStep(); // 90

            set.smsWorkbook = textBoxFileName.Text;

            set.Save();
            set.Reload();
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            // VALIDATE FILE EXISTANCE
            if ("".Equals(textBoxFileName.Text.Trim())) 
            {
                MessageBox.Show("No se ha seleccionado un archivo de excel para la configuración", "Archivo de Excel!");
                return;
            }

            if (!File.Exists(textBoxFileName.Text))
            {
                MessageBox.Show("No se encontró el archivo seleccionado", "Archivo de Excel!");
                return;
            }

            toolStripStatusLabel1.Text = "Validando libro seleccionado ...";
            toolStripProgressBar1.Visible = true;
            toolStripProgressBar1.PerformStep(); // 10
            // VALIDATE FILE CONTENTS
            ConfigValidator cv = new ConfigValidator();
            cv.API = api;
            toolStripProgressBar1.PerformStep(); // 20
            cv.Ruta = textBoxFileName.Text;
            toolStripProgressBar1.PerformStep(); // 30

            if (!cv.ValidateConfiguration())
            {   
                toolStripProgressBar1.Visible = false;
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Listo";

                MessageBox.Show("El archivo seleccionado no contiene información sobre la configuración de mensajes para metas de ventas", "Archivo de Excel!");
                return;
            }

            toolStripStatusLabel1.Text = "Grabando configuración ...";
            toolStripProgressBar1.Value = 100; // 100

            saveSettings();

            toolStripProgressBar1.Visible = false;

            System.Diagnostics.Process.Start(textBoxFileName.Text);
            this.Close();
        }
    }
}
