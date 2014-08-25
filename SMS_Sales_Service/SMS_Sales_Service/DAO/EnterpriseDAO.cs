using System.Collections.Generic;
using System.Data;
using Npgsql;

namespace SMSSender.DAO
{
    class EnterpriseDAO
    {
        public bool ErrorThrown = false;
        public string ErrorMessage = "";
        public List<Enterprise> listEnterprises = new List<Enterprise>();
        public List<Enterprise> readAll(NpgsqlConnection conn)
        {
            List<Enterprise> result = new List<Enterprise>();

            string sql = "SELECT distinct id_empresa, empresa " +
                            "FROM dim_sellers";
            
            DataTable dt = new DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql,conn);
            
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Enterprise enterprise = new Enterprise();

                    enterprise.Id = long.Parse(dt.Rows[i][0].ToString());
                    enterprise.Nombre = dt.Rows[i][1].ToString();

                    result.Add(enterprise);
                }
            }

            da.Dispose();
            listEnterprises = result;
            return result;
        }

        public Enterprise getEnterpriseByID(long id){
            return listEnterprises.Find(internalEnterprise => internalEnterprise.Id == id);
        }

    }
}
