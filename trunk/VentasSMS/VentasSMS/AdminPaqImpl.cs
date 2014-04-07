using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VentasSMS.AdminPaq.dto;

namespace VentasSMS
{
    public class AdminPaqImpl
    {
        private IList<Empresa> empresas = new List<Empresa>();
        private AdminPaqLib lib;

        public IList<Empresa> Empresas { get { return empresas; } set { empresas = value; } }
        public void addEmpresa(Empresa empresa) 
        {
            empresas.Add(empresa);
        }
        public void removeEmpresa(Empresa empresa)
        {
            empresas.Remove(empresa);
        }

        public AdminPaqImpl() {
            lib = new AdminPaqLib();
            lib.SetDllFolder();
        }
        public void InitializeSDK() {
            int connEmpresas, dbResponse, fieldResponse;
            connEmpresas = AdminPaqLib.dbLogIn("", lib.DataDirectory);

            if (connEmpresas == 0)
            {
                ErrLogger.Log("No se pudo crear conexión a la tabla de Empresas de AdminPAQ.");
                return;
            }

            dbResponse = AdminPaqLib.dbGetTopNoLock(connEmpresas, TableNames.EMPRESAS, IndexNames.EMPRESAS_PK);
            while (dbResponse == 0)
            {
                Empresa empresa = new Empresa();
                
                int idEmpresa = 0;
                fieldResponse = AdminPaqLib.dbFieldLong(connEmpresas, TableNames.EMPRESAS, Empresa.ID_EMPRESA, ref idEmpresa);
                empresa.Id = idEmpresa;

                StringBuilder nombreEmpresa = new StringBuilder(151);
                fieldResponse = AdminPaqLib.dbFieldChar(connEmpresas, TableNames.EMPRESAS, Empresa.NOMBRE_EMPRESA, nombreEmpresa, 151);
                string sNombreEmpresa = nombreEmpresa.ToString(0, 150).Trim();
                empresa.Nombre = sNombreEmpresa;

                StringBuilder rutaEmpresa = new StringBuilder(254);
                fieldResponse = AdminPaqLib.dbFieldChar(connEmpresas, TableNames.EMPRESAS, Empresa.RUTA_EMPRESA, rutaEmpresa, 254);
                string sRutaEmpresa = rutaEmpresa.ToString(0, 253).Trim();
                empresa.Ruta = sRutaEmpresa;

                empresas.Add(empresa);
                dbResponse = AdminPaqLib.dbSkip(connEmpresas, TableNames.EMPRESAS, IndexNames.EMPRESAS_PK, 1);
            }

            AdminPaqLib.dbLogOut(connEmpresas);
        }
    }
}
