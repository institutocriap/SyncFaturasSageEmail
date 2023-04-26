using SMSbyMail.SMSbyMailWS;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace notificacaoSemanalTestes
{
    class obj_logSend
    {
        public string colaboradorID { get; set; }
        public string mensagem { get; set; }
        public string nome { get; set; }
        public string numero { get; set; }

        public List<recipientWithName> smsSenders = new List<recipientWithName>();
    }
}
