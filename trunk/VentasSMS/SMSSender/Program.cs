using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using CommonAdminPaq;

namespace SMSSender
{
    class Program
    {
        static void Main(string[] args)
        {
            //Excel mode
            //if (args.Length == 0)
            //{
            //    ErrLogger.Log("USAGE - Unsuficient parameters, expectig to have a workbook path.");
            //    return;
            //}

            //SalesPicker sp = new SalesPicker();
            //sp.WorkbookPath = args[0];
            //sp.SendMethod = CommonConstants.SMS_LOCAL;
            //sp.UserName = "user";
            //sp.Password = "password";
            //sp.SendMethod = CommonConstants.SMS_MAS_MENSAJES;
            //sp.RetrieveConfiguration();
            //sp.RetrieveSales();
            //sp.SendSMS();
            //sp.SendBossSMS();
            //Console.WriteLine("*******************************************");
            //Console.ReadLine();



            //Postgresql mode
            SalesPickerPostgresql sp = new SalesPickerPostgresql();
            sp.UserName = "user";
            sp.Password = "password";
            sp.SendMethod = CommonConstants.SMS_MAS_MENSAJES;
            sp.RetrieveData();
            sp.SendSMS();
            sp.SendBossSMS();
            Console.WriteLine("*******************************************");
            Console.ReadLine();
        }
    }
}
