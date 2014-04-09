using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SMSSender
{
    interface ISendable
    {
        void SendSMS(Sms sms);
        void DoLogin(string userName, string password);
    }

    public class Sms {
        private string to;
        private string from;
        private string message;

        public string To { get { return to; } set { to = value; } }
        public string From { get { return from; } set { from = value; } }

        public string Message { get { return message; } 
            set 
            {
                value = removeTildes(value);
                if (value.Length > 150)
                    value = value.Substring(0, 150);

                message = value;
            } }

        private string removeTildes(string msg)
        {
            msg = msg.Replace('á', 'a');
            msg = msg.Replace('é', 'e');
            msg = msg.Replace('í', 'i');
            msg = msg.Replace('ó', 'o');
            msg = msg.Replace('ú', 'u');
            return msg;
        }
    }
}
