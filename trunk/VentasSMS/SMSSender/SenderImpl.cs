using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using CommonAdminPaq;

namespace SMSSender
{
    public class MasMensajes : ISendable
    {
        private const string API_URL = "https://www.masmensajes.com.mx/wss/smsapi13.php";
        private const string REQUEST = "{0}?usuario={1}&password={2}&celular={3}&mensaje={4}";

        private string userName, password;

        public void SendSMS(Sms sms)
        {
            string urlRequest;

            if (sms.Message.Length > 150)
                sms.Message = sms.Message.Substring(0, 150);

            urlRequest = string.Format(REQUEST, API_URL, userName, password, sms.To, Uri.EscapeUriString(sms.Message));
            WebRequest request = WebRequest.Create(urlRequest);

            Stream oStream = request.GetResponse().GetResponseStream();
            StreamReader objReader = new StreamReader(oStream);

            string sLine = "";
            int i = 0;

            while (sLine != null)
            {
                i++;
                sLine = objReader.ReadLine();
                if (sLine != null && sLine.Contains("ERROR"))
                    ErrLogger.Log(sLine);
            }
        }

        public void DoLogin(string userName, string password)
        {
            this.userName = userName;
            this.password = password;
        }
    }
    public class LocalSMS : ISendable
    {

        public void SendSMS(Sms sms)
        {
            string fileName = @"c:\temp\sms_" + sms.To + ".txt";
            using (StreamWriter w = File.AppendText(fileName))
            {
                w.WriteLine(sms.Message);
            }
        }


        public void DoLogin(string userName, string password)
        {
            Console.WriteLine(string.Format("Login - {0}: {1}", userName, password));
        }
    }
}
