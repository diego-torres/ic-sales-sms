﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonAdminPaq.dto;

namespace SMSSender.Entities
{
    class Director
    {
        public long ID;
        public bool SMS = false;
        public string Name;
        public string Email;
        public string CellPhone;
        public long empresa_id;

        public Enterprise empresa =null;
    }
}
