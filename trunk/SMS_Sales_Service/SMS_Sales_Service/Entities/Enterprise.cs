using SMSSender.Entities;
using System.Collections.Generic;


public class Enterprise
    {
        // FIELD INDEX IN TABLE
        public const int ID_EMPRESA = 1;
        public const int NOMBRE_EMPRESA = 2;
        public const int RUTA_EMPRESA = 3;

        // FIELD PROPERTIES FOR OBJECT
        private long id;
        private string nombre, ruta, alias, resultadoSemanal, resultadoMensual, resultadoDiario;
        private List<Seller> agentes = new List<Seller>();
        private List<Director> directors = new List<Director>();
        private IList<long> cSale = new List<long>();
        private IList<long> cReturn = new List<long>();
        

        public long Id { get { return id; } set { id = value; } }
        public string Nombre { get { return nombre; } set { nombre = value; } }
        public string Ruta { get { return ruta; } set { ruta = value; } }
        public string Alias { get { return alias; } set { alias = value; } }
        public List<Seller> Agentes { get { return agentes; } set { agentes = value; } }
        public IList<long> CodigosVenta { get { return cSale; } set { cSale = value; } }
        public IList<long> CodigosDevolucion { get { return cReturn; } set { cReturn = value; } }
        public List<Director> Directors { get { return directors; } set { directors = value; } }
        public string ResultadoSemanal { get { return resultadoSemanal; } set { resultadoSemanal = value; } }
        public string ResultadoMensual { get { return resultadoMensual; } set { resultadoMensual = value; } }
        public string ResultadoDiario { get { return resultadoDiario; } set { resultadoDiario = value; } }
    }
