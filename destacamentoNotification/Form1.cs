using criapLibrary.types;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using SyncFaturasSageEmail.Connects;
using System.Linq;
using System.Reflection;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using MailKit;
using System.IO;
using System.Text.RegularExpressions;

namespace SyncFaturasSageEmail
{
    public partial class Form1 : Form
    {
        public bool erro = false;
        public bool teste;
        Version versao = Assembly.GetExecutingAssembly().GetName().Version;
        string controleVersao = "";

        public Form1()
        {
            teste = true;

            InitializeComponent();
            Security.remote();
            controleVersao = @"Controle de versão: " + " V." + versao.Major.ToString() + "." + versao.Minor.ToString() + "." + versao.Build.ToString() + " Data de compilação do assembly: " + System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location) + " por rc";
            if (DateTime.Now.Hour == 22 && DateTime.Now.Minute <= 2)
            {
                EnviarRelatorioDiario();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                using (var cliente = new ImapClient())
                {
                    cliente.Connect("server.criap.com", 993, true);
                    cliente.Authenticate("docs-noreply@criap.com", "cqt6Y7uV{&ra");

                    var caixaEntrada = cliente.Inbox;
                    caixaEntrada.Open(MailKit.FolderAccess.ReadWrite);
                    var resultados = caixaEntrada.Search(SearchQuery.NotSeen);

                    foreach (var uid in resultados)
                    {
                        var mensagem = caixaEntrada.GetMessage(uid);

                        if (mensagem.Subject != null && mensagem.Subject.Contains("FATURA"))
                        {
                            if (mensagem.TextBody != null)
                            {
                                var regex = new Regex(@", Nº (\d+),");
                                var match = regex.Match(mensagem.TextBody);

                                if (match.Success)
                                {
                                    var numeroCapturado = match.Groups[1].Value;

                                    foreach (var anexo in mensagem.Attachments)
                                    {
                                        if (anexo is MimePart mimePart)
                                        {
                                            var nomeArquivoOriginal = mimePart.FileName;
                                            var partesNomeArquivo = nomeArquivoOriginal.Split('_');

                                            if (partesNomeArquivo.Length >= 3 && partesNomeArquivo[3] == numeroCapturado)
                                            {
                                                // Salva na base de dados
                                                var nomeArquivoOriginalSplit = nomeArquivoOriginal.Split('.');
                                                SalvaArquivoBasePA(mimePart, partesNomeArquivo, nomeArquivoOriginalSplit);
                                                DataBaseLogSaveSecretaria(partesNomeArquivo, $"Arquivo {nomeArquivoOriginal} salvo com sucesso.");
                                            }
                                            else
                                            {
                                                DataBaseLogSaveSecretaria(partesNomeArquivo, $"Número no nome do arquivo ({partesNomeArquivo[3]}) não corresponde ao número capturado ({numeroCapturado}).");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        // Marca o email como lido
                        caixaEntrada.AddFlags(uid, MessageFlags.Seen, true);
                    }
                    cliente.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                EnviarEmailErro(ex);
            }
        }

        private void SalvaArquivoBasePA(MimePart mimePart, string[] partesNomeArquivo, string[] nomeArquivoOriginal)
        {
            using (var memoriaStream = new MemoryStream())
            {
                mimePart.Content.DecodeTo(memoriaStream);
                byte[] conteudoArquivo = memoriaStream.ToArray();

                string selectQuery = @"SELECT COUNT(*) FROM TBForDocFaturas WHERE nome_arquivo = @nome_arquivo";

                Connect.SVlocalConnect.ConnInit();
                using (SqlCommand selectCommand = new SqlCommand(selectQuery, Connect.SVlocalConnect.Conn))
                {
                    selectCommand.Parameters.Add("@nome_arquivo", SqlDbType.NVarChar).Value = nomeArquivoOriginal[0];

                    int count = (int)selectCommand.ExecuteScalar();

                    string query;

                    if (count > 0)
                    {
                        query = @"UPDATE TBForDocFaturas 
                          SET ref_sage = @ref_sage, 
                              arquivo = @arquivo, 
                              extensao_arquivo = @extensao_arquivo, 
                              versao_data = @versao_data, 
                              versao_login = @versao_login
                          WHERE nome_arquivo = @nome_arquivo";
                    }
                    else
                    {
                        query = @"INSERT INTO TBForDocFaturas (ref_sage, arquivo, nome_arquivo, extensao_arquivo, versao_data, versao_login)
                          VALUES (
                              @ref_sage, @arquivo, @nome_arquivo, @extensao_arquivo, @versao_data, @versao_login
                          )";
                    }

                    using (SqlCommand comando = new SqlCommand(query, Connect.SVlocalConnect.Conn))
                    {
                        comando.Parameters.Add("@ref_sage", SqlDbType.NVarChar).Value = partesNomeArquivo[1] + " " + partesNomeArquivo[2] + "/" + partesNomeArquivo[3];
                        comando.Parameters.Add("@arquivo", SqlDbType.VarBinary).Value = conteudoArquivo;
                        comando.Parameters.Add("@nome_arquivo", SqlDbType.NVarChar).Value = nomeArquivoOriginal[0];
                        comando.Parameters.Add("@extensao_arquivo", SqlDbType.NVarChar).Value = nomeArquivoOriginal[1];
                        comando.Parameters.Add("@versao_data", SqlDbType.DateTime).Value = DateTime.Now;
                        comando.Parameters.Add("@versao_login", SqlDbType.NVarChar).Value = "system_auto_email_docs_noreply";

                        comando.ExecuteNonQuery();
                    }
                }
                Connect.SVlocalConnect.ConnEnd();
                Connect.closeAll();
            }
        }

        public string[] ObterFormandoId(string numDocPag)
        {
            string subQuery = $@"SELECT DISTINCT tf.Codigo_Formando, ta.Ref_Accao FROM TBForFormandos tf 
                        LEFT JOIN TBForFinOrdensFaturacaoPlano tp ON tp.Rowid_entidade = tf.versao_rowid 
                        LEFT JOIN TBForAccoes ta ON ta.versao_rowid = tp.Rowid_Opcao 
                        WHERE tp.num_doc_pag='{numDocPag}'";

            Connect.SVlocalConnect.ConnInit();
            SqlDataAdapter adapter = new SqlDataAdapter(subQuery, Connect.HTlocalConnect.Conn);
            DataTable subData = new DataTable();
            adapter.Fill(subData);
            Connect.SVlocalConnect.ConnEnd();
            Connect.closeAll();

            if (subData.Rows.Count > 0)
            {
                string[] retorno = new string[2];
                retorno[0] = subData.Rows[0]["Codigo_Formando"].ToString();
                retorno[1] = subData.Rows[0]["Ref_Accao"].ToString();
                return retorno;
            }
            return new string[] { "0", "0" };
        }

        private void DataBaseLogSaveSecretaria(string[] partesNomeArquivo, string mensagem)
        {
            string[] resultado = ObterFormandoId(partesNomeArquivo[1] + " " + partesNomeArquivo[2] + "/" + partesNomeArquivo[3]);
            string codigoFormando = resultado[0];
            string acao = resultado[1];
            string subQuery = $@"INSERT INTO sv_logs (idFormando, refAcao, dataregisto, registo, menu, username) VALUES ({resultado[0]}, '{resultado[1]}', GETDATE(), '{mensagem}', 'Faturas PA', 'system_auto_email_docs_noreply') ";

            Connect.SVlocalConnect.ConnInit();
            SqlCommand cmd = new SqlCommand(subQuery, Connect.SVlocalConnect.Conn);
            cmd.ExecuteNonQuery();
            Connect.SVlocalConnect.ConnEnd();
            Connect.closeAll();
        }

        private string GerarRelatorioDiario()
        {
            StringBuilder relatorio = new StringBuilder();
            relatorio.AppendLine("Relatório Diário de Arquivos Salvos");
            relatorio.AppendLine("Data: " + DateTime.Now.ToString("dd/MM/yyyy"));
            relatorio.AppendLine();
            relatorio.AppendLine("ID do Arquivo | Referência Sage | Data");

            string query = @"SELECT row_id, ref_sage, versao_data
                            FROM TBForDocFaturas
                            WHERE versao_data >= DATEADD(HOUR, 22, CAST(CONVERT(date, GETDATE() - 1) AS datetime))
                            AND versao_data < DATEADD(HOUR, 22, CAST(CONVERT(date, GETDATE()) AS datetime))";

            Connect.SVlocalConnect.ConnInit();
            using (SqlCommand comando = new SqlCommand(query, Connect.SVlocalConnect.Conn))
            using (SqlDataReader leitor = comando.ExecuteReader())
            {
                while (leitor.Read())
                {
                    string row_id = leitor["row_id"].ToString();
                    string refSage = leitor["ref_sage"].ToString();
                    string versao_data = leitor["versao_data"].ToString();
                    relatorio.AppendLine($"{row_id} | {refSage} | {versao_data}");
                }
            }
            Connect.SVlocalConnect.ConnEnd();
            Connect.closeAll();

            return relatorio.ToString();
        }

        private void EnviarRelatorioDiario()
        {
            string relatorio = GerarRelatorioDiario();
            var fromAddress = new MailAddress(Properties.Settings.Default.emailenvio, "Instituto CRIAP");
            var toAddress = new MailAddress("informatica@criap.com", "Equipe Técnica");
            string subject = "Rotina Importar Faturas Email || Relatório Diário de Arquivos Salvos";
            string body = $"Segue o relatório diário de arquivos salvos:\n\n{relatorio}\n\n{controleVersao}";

            var smtp = new SmtpClient
            {
                Host = "mail.criap.com",
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(Properties.Settings.Default.emailenvio, Properties.Settings.Default.passwordemail)
            };

            using (var mensagem = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(mensagem);
            }
        }

        private void EnviarEmailErro(Exception ex)
        {
            var fromAddress = new MailAddress(Properties.Settings.Default.emailenvio, "Instituto CRIAP");
            var toAddress = new MailAddress("informatica@criap.com", "Equipe Técnica");
            string subject = "Rotina Importar Faturas Email || Erro na Rotina";
            string body = $"Ocorreu um erro na rotina de verificação de emails:\n\n{ex.Message}\n\n{ex.StackTrace}\n\n{controleVersao}";

            var smtp = new SmtpClient
            {
                Host = "mail.criap.com",
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Credentials = new NetworkCredential(Properties.Settings.Default.emailenvio, Properties.Settings.Default.passwordemail)
            };

            using (var mensagem = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = body
            })
            {
                smtp.Send(mensagem);
            }
        }

    }
}
