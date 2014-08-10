using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonAdminPaq.dto;

namespace SMSSender.Entities
{
    class Seller
    {
        public long ID;
        public bool SMS = false;
        public long AP_ID = -1;
        public string Code = "";
        public string Name;
        public string Email;
        public string CellPhone;
        public float WeeklyGoal = 0;
        public float CumplimientoSemana = 0;
        public float CumplimientoMensual = 0;
        public long Enterprise_ID = -1;

        public Enterprise Empresa = null;
    }
}
