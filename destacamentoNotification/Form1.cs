using criapLibrary.types;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using notificacaoSemanalTestes.Connects;
using System.Linq;
using System.Reflection;

namespace notificacaoSemanalTestes
{
    public partial class Form1 : Form
    {
        public static List<FeriadosHT> feriadosHT = new List<FeriadosHT>();
        public List<objCalendario.modulo> modulosHT = new List<objCalendario.modulo>();
        public string assunto = "";
        public bool error = false;
        public bool teste;
        Version v = Assembly.GetExecutingAssembly().GetName().Version;
        string controloVersao = "";
        public class FeriadosHT
        {
            public DateTime Data { get; set; }
        }
        public Form1()
        {
            teste = true;
            InitializeComponent();
            Security.remote();
            controloVersao = @"<br><font size=""-2"">Controlo de versão: " + " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString() + " Assembly built date: " + System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location) + " by sa";
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                GetFeriadosHT();
            DateTime hoje, quarta1, quarta2;
            if (!teste)
            {
                hoje = DateTime.Now;
                //hoje = DateTime.Parse("2023-04-24");
                quarta1 = hoje.AddDays(2);
                quarta2 = quarta1.AddDays(7);
            }
            else
            {
                //hoje = DateTime.Now;
               hoje = DateTime.Parse("2024-06-10");
                //quarta1 = DateTime.Parse("2023-09-27");
                //quarta2 = DateTime.Parse("2023-10-04");
                quarta1 = hoje.AddDays(2);
                quarta2 = quarta1.AddDays(7);
            }
            string subQuery = "SELECT cav.codigo_curso, cav.RefAcao, cav.momentoav, cav.Metodologia, cav.moduloID, cast(cav.DataI_EN as date) as 'DataI_EN', cast(cav.Dataf_EN as date) as 'DataF_En' FROM CalendarioAV cav inner join humantrain.dbo.TBForCursos c on c.Codigo_Curso = cav.codigo_curso inner join humantrain.dbo.TBForModulos m on m.versao_rowid = cav.ModuloID inner join humantrain.dbo.TBForAccoes a on a.Ref_Accao = RefAcao where Metodologia not like '%atividade%' and DataI_EN >= '" + quarta1.ToString("yyyy-MM-dd") + " 00:00:00' and DAtaI_EN <= '" + quarta2.ToString("yyyy-MM-dd") + " 00:00:00' and a.Codigo_Estado = 1 group by cav.RefAcao, cav.MomentoAv, cav.Metodologia, cav.DataI_EN, cav.DataF_EN, cav.moduloid, cav.Codigo_Curso order by DataI_EN asc";


            Connect.SVlocalConnect.ConnInit();
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connect.SVlocalConnect.Conn);
            DataTable subData = new DataTable();
            adapter.Fill(subData);
            Connect.SVlocalConnect.ConnEnd();
            Connect.closeAll();

            string periodo = quarta1.ToLongDateString() + " a " + quarta2.ToLongDateString();
            string listaAcooes = "";
            assunto = "Testes de Avaliação || Período de " + periodo;

            listaAcooes = "Caros Colegas, <br>Informamos que para o período de <strong>" + periodo + "</strong> estão previstos, nos calendários de avaliação das ações que se seguem, os seguintes momentos de avaliação:<br><br>";

            string style = "style='color:white; background-color: #757271'";
            listaAcooes += " <table width=100% border=1 style='border-width: thin; border-collapse: collapse'><tr> <th colspan=8 " + style + " > <h3> " + assunto + " </h3> </th> </tr> " +
            " <th " + style + " >Ação de Formação</th>" +
            " <th " + style + " >Coordenador Científico / Técnico de Formação</th>" +
            " <th " + style + " >Técnico EAD</th>" +
            " <th " + style + " >Momento de Avaliação</th>" +
            " <th " + style + " >Módulo / Arquivo</th> " +
            " <th " + style + " >Metodologia</th> " +
            " <th " + style + " >Data de início</th> " +
            " <th " + style + " >Data de fim</th> ";

            if (subData.Rows.Count > 0)
            {
                for (int i = 0; i < subData.Rows.Count; i++)
                {
                    DataRow row = subData.Rows[i];
                    string refacao = row["RefAcao"].ToString();
                    string codcurso = row["codigo_curso"].ToString();
                    string momAv = row["momentoav"].ToString().Split(' ')[1];
                    string metodologia = row["Metodologia"].ToString();
                    string dt1 = DateTime.Parse(row["DataI_EN"].ToString()).ToString("dd/MM/yyyy");
                    DateTime dfAux;
                    string dt2 = "";
                    if (DateTime.TryParse(row["DataF_En"].ToString(), out dfAux))
                    {
                        dt2 = dfAux.ToString("dd/MM/yyyy");
                    }
                    else
                    {
                        dt2 = dt1;
                    }

                    string modulosText = "", tecnRespon = "", TecnAssis = "";
                    int modid = int.Parse(row["moduloID"].ToString());

                    GetModulosHT(codcurso, refacao);

                    if (modulosHT.Count > 0)
                        try
                        {
                            modulosText = modulosHT.Where(x => x.Modulo_ID == modid).First().Descricao;
                        }
                        catch
                        {
                            modulosText = "Sem descrição";
                        }

                    string strtecnAss = "SELECT sy.Nome as 'TecResp', sy.Login, sx.Nome as 'TecAssist' FROM TBForAccoes a left join TBSysUsers sy on a.Cod_Tecn_Resp = sy.Login left join TBSysUsers sx on a.Cod_Tecn_Assist = sx.Login where Ref_Accao ='" + refacao + "'";
                    Connect.HTlocalConnect.ConnInit();
                    SqlDataAdapter adptrtecass = new SqlDataAdapter(strtecnAss, Connect.HTlocalConnect.Conn);
                    DataTable subDtecass = new DataTable();
                    adptrtecass.Fill(subDtecass);
                    Connect.HTlocalConnect.ConnEnd();
                    if (subDtecass.Rows.Count > 0)
                    {
                            tecnRespon = subDtecass.Rows[0]["TecResp"]?.ToString() ?? "n/a";

                            TecnAssis = subDtecass.Rows[0]["TecAssist"]?.ToString() == "" ? "n/a" : subDtecass.Rows[0]["TecAssist"].ToString() ?? "n/a";

                        }
                        //else continue;

                        listaAcooes += " <tr style='font-size: 13px; background-color: white'> " + " <td align=center> " + refacao + "</td> <td align=center> " + tecnRespon + "</td> <td align=center> " + TecnAssis + "</td> <td align=center> " + momAv + "</td> <td align=center> " + modulosText + "</td> <td align=center> " + metodologia + "</td> <td align=center> " + dt1 + "</td> <td align=center> " + dt2 + "</td>";
                }
            }


            string subQuery2 = "SELECT a.Ref_Accao, a.Data_Fim, c.Tipo_Curso, cpa.PARAM_AVAL, cpa.DESCRICAO from TBForAccoes a inner join TBForCursos c on a.Codigo_Curso = c.Codigo_Curso inner join TBForCursosParamsAvalQtva cpa on c.Codigo_Curso = cpa.Codigo_Curso left join [secretariaVirtual].[dbo].[CalendarioAV] cav on a.Ref_Accao=cav.RefAcao where a.Ref_Accao is not null and c.Tipo_Curso = '2' and a.Data_Fim between dateadd(day, -14, '" + hoje.ToString("yyyy-MM-dd") + "') and '" + hoje.ToString("yyyy-MM-dd") + "' and Momento_Aval = 1 and Nivel_Modulo = 0 and a.Codigo_Estado=1 and cav.RefAcao is null order by a.Data_Fim asc";

            Connect.HTlocalConnect.ConnInit();
            SqlDataAdapter adapter2 = new SqlDataAdapter(subQuery2, Connect.HTlocalConnect.Conn);
            DataTable subData2 = new DataTable();
            adapter2.Fill(subData2);
            Connect.HTlocalConnect.ConnEnd();
            Connect.closeAll();
            if (subData2.Rows.Count > 0)
            {
                listaAcooes += " </table>  <br><br><strong>Ações que possivelmente não têm Calendário de Avaliação (CA) criado na SV:</strong><br> <br>";
                style = "style='color:white; background-color: #D1030B'";
                listaAcooes += " <table width=100% border=1 style='border-width: thin; border-collapse: collapse'><tr> <th colspan=8 " + style + " > <h3> " + assunto + " </h3> </th> </tr> " +
            " <th " + style + " >Ação de Formação</th>" +
            " <th " + style + " >Coordenador Científico / Técnico de Formação</th>" +
            " <th " + style + " >Técnico EAD</th>" +
            " <th " + style + " >Momento de Avaliação</th>" +
            " <th " + style + " >Módulo / Arquivo</th> " +
            " <th " + style + " >Metodologia</th> " +
            " <th " + style + " >Data de início</th> " +
            " <th " + style + " >Data de fim</th> ";
                for (int i = 0; i < subData2.Rows.Count; i++)
                {
                    DataRow row = subData2.Rows[i];
                    DateTime datainicio = DataCalendarioIN(DateTime.Parse(row["Data_Fim"].ToString()));
                    string refacao = row["Ref_Accao"].ToString();

                    if (!listaAcooes.Contains(refacao))
                    {
                        if (datainicio >= quarta1 && datainicio <= quarta2)
                        {
                            DateTime datafim = DataCalendarioFim(datainicio);
                            string tecnRespon = "", TecnAssis = "";

                            string metodologia = row["DESCRICAO"].ToString();
                            string strtecnAss = "SELECT sy.Nome as 'TecResp', sy.Login, sx.Nome as 'TecAssist' FROM TBForAccoes a left join TBSysUsers sy on a.Cod_Tecn_Resp = sy.Login left join TBSysUsers sx on a.Cod_Tecn_Assist = sx.Login where Ref_Accao ='" + refacao + "'";
                            Connect.HTlocalConnect.ConnInit();
                            SqlDataAdapter adptrtecass = new SqlDataAdapter(strtecnAss, Connect.HTlocalConnect.Conn);
                            DataTable subDtecass = new DataTable();
                            adptrtecass.Fill(subDtecass);
                            Connect.HTlocalConnect.ConnEnd();
                            if (subDtecass.Rows.Count > 0)
                            {
                                tecnRespon = subDtecass.Rows[0]["TecResp"]?.ToString()??"n/a";
                                TecnAssis = subDtecass.Rows[0]["TecAssist"]?.ToString() == "" ? "n/a" : subDtecass.Rows[0]["TecAssist"].ToString() ?? "n/a";

                                }
                                else continue;

                            listaAcooes += " <tr style='font-size: 13px; background-color: white'> " +
                            " <td align=center> " + refacao
                            + "</td> <td align=center> " + tecnRespon
                            + "</td> <td align=center> " + TecnAssis
                            + "</td> <td align=center> " + "I"
                            + "</td> <td align=center> " + " "
                            + "</td> <td align=center> " + metodologia
                            + "</td> <td align=center> " + datainicio.ToString("dd/MM/yyyy")
                            + "</td> <td align=center> " + datafim.ToString("dd/MM/yyyy") + "</td>";
                        }
                    }
                }
            }

            listaAcooes += " </table>  <br><br>Agradecemos a validação da conformidade da informação apresentada, <br>Instituto CRIAP<br>" + controloVersao; //responsable
            SendLog(listaAcooes);
            }
            catch (Exception ex)
            {
                //LogWrite("\tERROR\t" + ex.Message + "\r\n");
                error = true;
                SendLog("\tERROR\t" + ex.Message + "\r\n" + controloVersao);
            }

            Connect.closeAll();
            Application.Exit();
        }
        private static DateTime DataCalendarioIN(DateTime data)
        {
            DateTime dataIn = DateTime.MinValue;
            if (data.DayOfWeek == DayOfWeek.Monday) dataIn = data.AddDays(13);
            else if (data.DayOfWeek == DayOfWeek.Tuesday) dataIn = data.AddDays(12);
            else if (data.DayOfWeek == DayOfWeek.Wednesday) dataIn = data.AddDays(11);
            else if (data.DayOfWeek == DayOfWeek.Thursday) dataIn = data.AddDays(10);
            else if (data.DayOfWeek == DayOfWeek.Friday) dataIn = data.AddDays(9);
            else if (data.DayOfWeek == DayOfWeek.Saturday) dataIn = data.AddDays(8);
            else if (data.DayOfWeek == DayOfWeek.Sunday) dataIn = data.AddDays(14);

            int existe = feriadosHT.Where(x => x.Data.Date == dataIn.Date).Count();
            if (existe > 0) dataIn = dataIn.AddDays(-1);
            return dataIn;
        }
        private static DateTime DataCalendarioFim(DateTime datain)
        {
            DateTime dataFim = datain.AddDays(1);
            int existe = feriadosHT.Where(x => x.Data.Date == dataFim.Date).Count();
            if (existe > 0) dataFim = dataFim.AddDays(1);
            return dataFim;
        }
        private static void GetFeriadosHT()
        {
            feriadosHT.Clear();
            string subQuery = "select Data from TBGerFeriados";
            Connect.HTlocalConnect.ConnInit();
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connect.HTlocalConnect.Conn);
            DataTable subData = new DataTable();
            adapter.Fill(subData);
            List<DataRow> dataList = subData.AsEnumerable().ToList();
            Connect.HTlocalConnect.ConnEnd();
            Connect.closeAll();
            if (dataList.Count > 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new FeriadosHT()
                    {
                        Data = DateTime.Parse(dataList[i][0].ToString()),
                    };
                    feriadosHT.Add(obj);
                }
            }
        }
        public void GetModulosHT(string codCurso, string refCurso)
        {
            modulosHT.Clear();
            List<DataRow> dataList = GetModulosHT(codCurso, refCurso, Connect.HTlocalConnect.Conn);

            if (dataList.Count != 0)
            {
                for (int i = 0; i < dataList.Count; i++)
                {
                    var obj = new objCalendario.modulo()
                    {
                        Nivel = dataList[i][1].ToString(),
                        Descricao = dataList[i][2].ToString(),
                        N_Horas = float.Parse(dataList[i][3].ToString()),
                        Peso_Aval = float.Parse(dataList[i][4].ToString()),
                        Data_Inicio = (DateTime)(dataList[i][6]),
                        Data_Fim = (DateTime)(dataList[i][7]),
                        Modulo_ID = (int)dataList[i][0],
                        Ordem = -1,
                        Teste_ID = -1,
                        check = false
                    };
                    modulosHT.Add(obj);
                }

                int j = 0;
                while (j < modulosHT.Count - 1)
                {
                    for (int k = j; k < modulosHT.Count - 1; k++)
                    {
                        if (modulosHT[j].Nivel == modulosHT[k + 1].Nivel)
                        {
                            modulosHT[j].Data_Inicio = modulosHT[k + 1].Data_Inicio;
                            modulosHT.RemoveAt(k + 1);
                            k--;
                        }
                    }
                    j++;
                }
            }
        }
        public static List<DataRow> GetModulosHT(string codCurso, string refCurso, SqlConnection connection)
        {
            SqlDataAdapter sqlDataAdapter = new SqlDataAdapter("SELECT m.versao_rowid, m.Nivel, m.Descricao, m.N_Horas, m.Peso_Aval, s.Data, s.Hora_Inicio, s.Hora_Fim, ROW_NUMBER() OVER(ORDER BY s.Hora_Fim DESC) as RowNum FROM TBForModulos m INNER JOIN TBForSessoes s ON m.versao_rowid = s.Rowid_Modulo INNER JOIN TBForAccoes a ON s.Rowid_Accao = a.versao_rowid WHERE(m.Codigo_Curso = '" + codCurso + "') AND(a.Ref_Accao = '" + refCurso + "') ", connection);
            DataTable dataTable = new DataTable();
            sqlDataAdapter.Fill(dataTable);
            List<DataRow> list = dataTable.AsEnumerable().ToList<DataRow>();
            connection.Close();
            return list;
        }
        public void LogWrite(string text)
        {
            richTextBox1.AppendText(DateTime.Now.ToString("dd-MM-yyyy H:mm") + text);
            richTextBox1.ScrollToCaret();
        }
        private void SendLog(string body)
        {
            NetworkCredential basicCredential = new NetworkCredential(Properties.Settings.Default.emailenvio, Properties.Settings.Default.passwordemail);

            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.Host = "mail.criap.com";
            client.Timeout = 10000;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Credentials = basicCredential;

            MailMessage mm = new MailMessage();
            mm.From = new MailAddress("Instituto CRIAP <" + Properties.Settings.Default.emailenvio + "> ");

            if (!error)
            {
                if (!teste)
                {
                    mm.To.Add("ead@criap.com");
                    mm.To.Add("formacao@criap.com");
                    mm.To.Add("informatica@criap.com");
                }
                else
                    mm.To.Add("sandraaguilar@criap.com");
            }
            else
            {
                if (!teste)
                    mm.To.Add("informatica@criap.com");
                else
                    mm.To.Add("sandraaguilar@criap.com");
                mm.Body = richTextBox1.Text;
                mm.Subject = "ERRO - ";
            }

            mm.Subject += assunto;
            mm.IsBodyHtml = false;
            mm.BodyEncoding = UTF8Encoding.UTF8;
            mm.IsBodyHtml = true;
            mm.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;
            mm.Body += body;

            client.Send(mm);
        }
    }
}