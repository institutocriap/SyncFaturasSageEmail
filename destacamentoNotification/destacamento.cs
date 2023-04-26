using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace notificacaoSemanalTestes
{
    class destacamento
    {
        public string rowid { get; set; }
        public string rowid_Sessao { get; set; }
        public string rowid_TipoDest { get; set; }
        public string codigo_Colaborador { get; set; }
        public DateTime data_Inicio { get; set; }
        public DateTime data_Fim { get; set; }
        public string descricao { get; set; }
        public string notas { get; set; }
        public string local { get; set; }
        public string local_Original { get; set; }
        public string local_Live { get; set; }
        public string codigo_Google { get; set; }
        public int label { get; set; }
        public int colaboradorID { get; set; }
        public string codigo_Validacao { get; set; }
    }
}
