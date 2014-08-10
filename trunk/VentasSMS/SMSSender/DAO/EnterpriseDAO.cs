using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMSSender.Entities;
using CommonAdminPaq.dto;
using System.Data;
using Npgsql;

namespace SMSSender.DAO
{
    class EnterpriseDAO
    {
        public bool ErrorThrown = false;
        public string ErrorMessage = "";
        public List<Enterprise> listEnterprises;
        public List<Enterprise> readAll(NpgsqlConnection conn)
        {
            List<Enterprise> result = null;

            string sql = "SELECT seller_id, sms, ap_id, agent_code, agent_name, email, cellphone, weekly_goal, id_empresa " +
                            "FROM dim_sellers where sms = true";
            
            DataTable dt = new DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql,conn);
            
            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                Enterprise enterprise= new Enterprise();
                
                enterprise.Id = long.Parse(dt.Rows[0][0].ToString());
                enterprise.Nombre = dt.Rows[0][1].ToString();
                enterprise.Ruta = dt.Rows[0][2].ToString();
                
                result.Add(enterprise);
            }

            da.Dispose();
            listEnterprises = result;
            return result;
        }

        public Enterprise getEnterpriseByID(long id){
            return listEnterprises.Find(internalEnterprise => internalEnterprise.Id == id);
        }

        public List<Empresa> readEmpresasInUse()
        {
            List<Empresa> result = null;

            foreach (Seller seller in listSellers)
            {
                Empresa currentEmpresa = seller.Empresa;
                
                Empresa alreadyExists = result.Find(internalEmpresa => internalEmpresa.Id == currentEmpresa.Id);
                if (alreadyExists == null)
                {
                    result.Add(currentEmpresa);
                }
            }

            return result;
        }
    }
}
