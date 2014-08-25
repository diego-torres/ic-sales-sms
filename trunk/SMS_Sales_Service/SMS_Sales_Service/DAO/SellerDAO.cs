using System.Collections.Generic;
using SMSSender.Entities;
using Npgsql;
using System.Data;

namespace SMSSender.DAO
{
    class SellerDAO
    {
        public bool ErrorThrown = false;
        public string ErrorMessage = "";
        public List<Seller> listSellers = new List<Seller>();
        public List<Seller> readAll(NpgsqlConnection conn)
        {
            List<Seller> result = new List<Seller>();

            string sql = "SELECT dim_sellers.seller_id,sms,ap_id,agent_code,agent_name,email,dim_sellers.cellphone,weekly_goal, " +
                            "id_empresa	,sum(fact_sales.sold_week) AS sold_week	,sum(fact_sales.sold_month) AS sold_month " +
                            "FROM dim_sellers LEFT JOIN fact_sales ON dim_sellers.seller_id = fact_sales.seller_id WHERE sms = true  " +
                            "GROUP BY dim_sellers.seller_id,sms,ap_id,agent_code,agent_name,email,dim_sellers.cellphone,weekly_goal,id_empresa";

            DataTable dt = new DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);

            da.Fill(dt);

            if (dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    Seller seller = new Seller();

                    seller.ID = long.Parse(dt.Rows[i][0].ToString());
                    seller.SMS = bool.Parse(dt.Rows[i][1].ToString());
                    seller.AP_ID = long.Parse(dt.Rows[i][2].ToString());
                    seller.Code = dt.Rows[i][3].ToString();
                    seller.Name = dt.Rows[i][4].ToString();
                    seller.Email = dt.Rows[i][5].ToString();
                    seller.CellPhone = dt.Rows[i][6].ToString();
                    seller.WeeklyGoal = float.Parse(dt.Rows[i][7].ToString());
                    seller.Enterprise_ID = long.Parse(dt.Rows[i][8].ToString());
                    if (dt.Rows[i][9].ToString() == "")
                    {
                        seller.CumplimientoSemana = 0;
                    }
                    else
                    {
                        seller.CumplimientoSemana = float.Parse(dt.Rows[i][9].ToString());
                    }
                    if (dt.Rows[i][10].ToString() == "")
                    {
                        seller.CumplimientoMensual = 0;
                    }
                    else
                    {
                        seller.CumplimientoMensual = float.Parse(dt.Rows[i][10].ToString());
                    }
                    
                    result.Add(seller);
                }
            }

            da.Dispose();

            listSellers = result;
            return result;
        }
        public void setEnterprisesToSellers(EnterpriseDAO enterpriseDAO)
        {
            List<Enterprise> allEnterprises = enterpriseDAO.listEnterprises;
            foreach (Seller seller in listSellers)
            {
                seller.Empresa = enterpriseDAO.getEnterpriseByID(seller.Enterprise_ID);
            }
        }
        public List<Enterprise> readEmpresasInUse(List<Enterprise> currentListEnterprises)
        {
            foreach (Seller seller in listSellers)
            {
                Enterprise currentEmpresa = seller.Empresa;

                if (currentEmpresa != null)
                {
                    Enterprise alreadyExists = currentListEnterprises.Find(internalEmpresa => internalEmpresa.Id == currentEmpresa.Id);
                    if (alreadyExists == null)
                    {
                        currentEmpresa.Agentes.Add(seller);
                        currentListEnterprises.Add(currentEmpresa);
                    }
                    else
                    {
                        alreadyExists.Agentes.Add(seller);
                    }
                }
            }

            return currentListEnterprises;
        }
    }
}
