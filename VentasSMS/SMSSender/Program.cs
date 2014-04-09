using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CommonAdminPaq;

namespace SMSSender
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                ErrLogger.Log("USAGE - Unsuficient parameters, expectig to have a workbook path.");
                return;
            }

            SalesPicker sp = new SalesPicker();
            sp.WorkbookPath = args[0];
            sp.SendMethod = CommonConstants.SMS_MAS_MENSAJES;
            sp.RetrieveConfiguration();
            sp.RetrieveSales();
            sp.SendSMS();
            sp.SendBossSMS();
            Console.ReadLine();
        }
    }
}
