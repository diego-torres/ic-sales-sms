using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace CommonAdminPaq
{
    public class ErrLogger
    {
        public static void Log(string message) {
            string fileName = @"c:\temp\err_" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            using (StreamWriter w = File.AppendText(fileName))
            {
                w.WriteLine(message);
            }
        }
    }
}
