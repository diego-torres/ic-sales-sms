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

        /// <summary>
        /// Validar hoja de excel
        /// </summary>
        /// <param name="evalSheet"></param>
        /// <returns></returns>
        private bool validateWorkSheet(Excel.Worksheet evalSheet)
        {
            toolStripStatusLabel1.Text = "Validando contenido de la hoja: " + evalSheet.Name;
            // Validate
            // Check company name
            string sEmpresa = evalSheet.Range["B2"].Value;
            bool empresaEnLista = false;
            foreach(Empresa empresa in api.Empresas){
                //sEmpresa.ToLower().Equals(empresa.Nombre.ToLower())
                if (sEmpresa.Equals(empresa.Nombre)) {
                    empresaEnLista = true;
                    evalSheet.Range["B1"].Value = empresa.Id;
                    evalSheet.Range["B3"].Value = empresa.Ruta;
                    break;
                }
            }
            if (!empresaEnLista)
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }

            // Validar codigos de documentos
            /*bool codigoEsValido = false;
              //buscar en api
              // si lo encuentras, codigoEsValido = true;
            if (!codigoEsValido)
            {
                evalSheet.Range["D1"].Value = "";
                return false;
            }*/

            // Validar codigos de agente y poner el sysid
            

            evalSheet.Range["D1"].Value = "sysvalidated";
            return true;
        }

        private bool validateFile(string fileName) 
        {
            bool validated = false;
            int sheetIdx = 1;
            // Validate each worksheet
            toolStripProgressBar1.PerformStep(); //20

            toolStripStatusLabel1.Text = "Entrando a librerías de Excel ...";
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = false;


            toolStripProgressBar1.PerformStep(); // 30
            toolStripProgressBar1.PerformStep(); // 40

            toolStripStatusLabel1.Text = "Cargando libro de excel en memoria ...";
            Excel.Workbook book = xlApp.Workbooks.Open(fileName);
            toolStripProgressBar1.PerformStep(); // 50
            toolStripProgressBar1.PerformStep(); // 60

            toolStripStatusLabel1.Text = "Validando hojas de configuración ...";
            while (sheetIdx <= book.Worksheets.Count) {
                validated = validateWorkSheet(book.Worksheets[sheetIdx]);
                if (!validated) return false;
                sheetIdx++;
            }

            toolStripStatusLabel1.Text = "Liberando recursos de excel ...";
            toolStripProgressBar1.PerformStep(); // 70

            if (validated) book.Save();

            book.Close();
            xlApp.Quit();

            Marshal.FinalReleaseComObject(book);
            Marshal.FinalReleaseComObject(xlApp);

            toolStripProgressBar1.PerformStep(); // 80
            
            return validated;
        }

        private void saveSettings()
        {
            toolStripProgressBar1.PerformStep(); // 90
            // Open App.Config of executable
            System.Configuration.Configuration config =
             ConfigurationManager.OpenExeConfiguration
                        (ConfigurationUserLevel.None);

            config.AppSettings.Settings.Remove(CommonConstants.SMS_FILE_NAME);
            //config.AppSettings.Settings.Clear();

            // Add an Application Setting.
            config.AppSettings.Settings.Add(CommonConstants.SMS_FILE_NAME, textBoxFileName.Text);

            // Save the changes in App.config file.
            config.Save(ConfigurationSaveMode.Modified);

            // Force a reload of a changed section.
            ConfigurationManager.RefreshSection("appSettings");
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
            if (!validateFile(textBoxFileName.Text))
            {
                toolStripProgressBar1.Visible = false;
                toolStripProgressBar1.Value = 0;
                toolStripStatusLabel1.Text = "Listo";

                MessageBox.Show("El archivo seleccionado no contiene información sobre la configuración de mensajes para metas de ventas", "Archivo de Excel!");
                return;
            }

            toolStripStatusLabel1.Text = "Grabando configuración ...";

            saveSettings();

            toolStripProgressBar1.PerformStep(); // 100
            toolStripProgressBar1.Visible = false;

            System.Diagnostics.Process.Start(textBoxFileName.Text);
            this.Close();
        }
    }
}
