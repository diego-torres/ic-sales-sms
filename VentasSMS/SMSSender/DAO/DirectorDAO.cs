using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SMSSender.Entities;
using CommonAdminPaq.dto;
using Npgsql;
using System.Data;

namespace SMSSender.DAO
{
    class DirectorDAO
    {
        public bool ErrorThrown = false;
        public string ErrorMessage = "";
        public List<Director> listDirectors = new List<Director>();
        public List<Director> readAll(NpgsqlConnection conn)
        {
            List<Director> result = new List<Director>();

            string sql = "SELECT director_id, sms, director_name, email, cellphone, id_empresa " +
                            "FROM dim_directors where sms = true";
            
            DataTable dt = new DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql,conn);
            
            da.Fill(dt);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Director director = new Director();

                director.ID = long.Parse(dt.Rows[i][0].ToString());
                director.SMS = bool.Parse(dt.Rows[i][1].ToString());
                director.Name = dt.Rows[i][2].ToString();
                director.Email = dt.Rows[i][3].ToString();
                director.CellPhone = dt.Rows[i][4].ToString();
                director.empresa_id = long.Parse(dt.Rows[i][5].ToString());

                result.Add(director);
            }
            
            da.Dispose();

            listDirectors = result;
            return result;
        }

        public void setEnterprisesToDirectors(EnterpriseDAO enterpriseDAO)
        {
            List<Enterprise> allEnterprises = enterpriseDAO.listEnterprises;
            foreach (Director director in listDirectors)
            {
                director.empresa = enterpriseDAO.getEnterpriseByID(director.empresa_id);
            }
        }
        public List<Enterprise> readEmpresasInUse(List<Enterprise> currentListEnterprises)
        {
            foreach (Director director in listDirectors)
            {
                Enterprise currentEmpresa = director.empresa;
                if (currentEmpresa != null)
                {
                    Enterprise alreadyExists = currentListEnterprises.Find(internalEmpresa => internalEmpresa.Id == currentEmpresa.Id);
                    if (alreadyExists == null)
                    {
                        currentEmpresa.Directors.Add(director);
                        currentListEnterprises.Add(currentEmpresa);
                    }
                    else
                    {
                        alreadyExists.Directors.Add(director);
                    }
                }
            }

            return currentListEnterprises;
        }
    }
}
