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
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using MailKit;
using System.IO;
using System.Text.RegularExpressions;

namespace notificacaoSemanalTestes
{
    public partial class Form1 : Form
    {
        
        public bool error = false;
        public bool teste;
        Version v = Assembly.GetExecutingAssembly().GetName().Version;
        string controloVersao = "";
       
        public Form1()
        {
            teste = true;
            InitializeComponent();
            Security.remote();
            controloVersao = @"<br><font size=""-2"">Controlo de versão: " + " V." + v.Major.ToString() + "." + v.Minor.ToString() + "." + v.Build.ToString() + " Assembly built date: " + System.IO.File.GetLastWriteTime(Assembly.GetExecutingAssembly().Location) + " by sa";
        }
        private void Form1_Load(object sender, EventArgs e)
        {

            using (var client = new ImapClient())
            {
                // Conecte-se ao servidor IMAP
                client.Connect("server.criap.com", 993, true);

                // Autenticação
                client.Authenticate("docs-noreply@criap.com", "cqt6Y7uV{&ra");

                // Seleciona a caixa de entrada
                var inbox = client.Inbox;
                inbox.Open(MailKit.FolderAccess.ReadWrite);

                // Buscar emails não lidos
                var results = inbox.Search(SearchQuery.NotSeen);

                foreach (var uid in results)
                {
                    var message = inbox.GetMessage(uid);

                    if (message.Subject != null && message.Subject.Contains("FATURA"))
                    {
                        if (message.TextBody != null)
                        {
                            // Expressão regular para capturar o número após "Nº" e entre vírgulas
                            var regex = new Regex(@", Nº (\d+),");
                            var match = regex.Match(message.TextBody);

                            if (match.Success)
                            {
                                var numeroCapturado = match.Groups[1].Value;

                                foreach (var attachment in message.Attachments)
                                {
                                    if (attachment is MimePart mimePart)
                                    {
                                        var originalFilename = mimePart.FileName;
                                        var fileNameParts = originalFilename.Split('_');

                                        if (fileNameParts.Length >= 3 && fileNameParts[3] == numeroCapturado)
                                        {
                                            // Salva local
                                            var filePath = Path.Combine(@"C:\Users\raphaelcastro\Downloads\", originalFilename);
                                            
                                            using (var stream = File.Create(filePath))
                                            {
                                                mimePart.Content.DecodeTo(stream);
                                            }

                                            // Salva base de dados
                                            SalvaArquivoBasePA(mimePart, originalFilename);

                                            DataBaseLogSaveSecretaria(fileNameParts, $"Arquivo {originalFilename} salvo com sucesso.");
                                        }
                                        else
                                        {
                                            DataBaseLogSaveSecretaria(fileNameParts, $"Número no nome do arquivo ({fileNameParts[3]}) não corresponde ao número capturado ({numeroCapturado}).");
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Marca o email como lido
                    inbox.AddFlags(uid, MessageFlags.Seen, true);
                }
                client.Disconnect(true);
            }
        }
        private void SalvaArquivoBasePA(MimePart mimePart, string nomeArquivoOriginal)
        {
            // Lê o conteúdo do arquivo diretamente do anexo
            using (var memoriaStream = new MemoryStream())
            {
                mimePart.Content.DecodeTo(memoriaStream);
                byte[] conteudoArquivo = memoriaStream.ToArray();

                string query = @"INSERT INTO TBForDocFaturas (FileName, FileContent) 
                         VALUES (@FileName, @FileContent)";

                Connect.portalAlunoConnect.ConnInit();
                using (SqlCommand comando = new SqlCommand(query, Connect.portalAlunoConnect.Conn))
                {
                    comando.Parameters.Add("@FileName", SqlDbType.NVarChar).Value = nomeArquivoOriginal;
                    comando.Parameters.Add("@FileContent", SqlDbType.VarBinary).Value = conteudoArquivo;
                    comando.ExecuteNonQuery();
                }
                Connect.SVlocalConnect.ConnEnd();
                Connect.closeAll();
            }
        }

        public string[] getFormandoId(string num_doc_pag)
        {
            string subQuery = $@"SELECT DISTINCT tf.Codigo_Formando, ta.Ref_Accao FROM TBForFormandos tf 
                        LEFT JOIN TBForFinOrdensFaturacaoPlano tp ON tp.Rowid_entidade = tf.versao_rowid 
                        LEFT JOIN TBForAccoes ta ON ta.versao_rowid = tp.Rowid_Opcao 
                        WHERE tp.num_doc_pag='{num_doc_pag}'";

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

        private void DataBaseLogSaveSecretaria(string[] fileNameParts, string mensagem)
        {
            // Busca o id formando
            string[] resultado = getFormandoId(fileNameParts[1] + " " + fileNameParts[2] + "/" + fileNameParts[3]);
            string codigoFormando = resultado[0];
            string acao = resultado[1];
            string subQuery = $@"INSERT INTO sv_logs (idFormando, refAcao, dataregisto, registo, menu, username) VALUES ({resultado[0]}, '{resultado[1]}', GETDATE(), '{mensagem}', 'Faturas PA', 'system') ";

            Connect.SVlocalConnect.ConnInit();
            SqlCommand cmd = new SqlCommand(subQuery, Connect.SVlocalConnect.Conn);
            cmd.ExecuteNonQuery();
            Connect.SVlocalConnect.ConnEnd();
            Connect.closeAll();
        }

    }
}