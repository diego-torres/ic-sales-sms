using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SMSSender
{
    public class MasMensajes : ISendable
    {

        public void SendSMS(Sms sms)
        {
            throw new NotImplementedException();
        }
    }
    public class LocalSMS : ISendable
    {

        public void SendSMS(Sms sms)
        {
            throw new NotImplementedException();
        }
    }
}
