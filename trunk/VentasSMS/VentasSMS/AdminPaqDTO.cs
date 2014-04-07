using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace VentasSMS.AdminPaq.dto
{
    public class TableNames {
        public const string EMPRESAS = "MGW00001";
    }
    public class IndexNames {
        public const string EMPRESAS_PK = "PRIMARYKEY";
    }
    public class Empresa
    {
        // FIELD INDEX IN TABLE
        public const int ID_EMPRESA = 1;
        public const int NOMBRE_EMPRESA = 2;
        public const int RUTA_EMPRESA = 3;

        // FIELD PROPERTIES FOR OBJECT
        private long id;
        private string nombre;
        private string ruta;
        public long Id { get { return id; } set { id = value; } }
        public string Nombre { get { return nombre; } set { nombre = value; } }
        public string Ruta { get { return ruta; } set { ruta = value; } }
    }
}
