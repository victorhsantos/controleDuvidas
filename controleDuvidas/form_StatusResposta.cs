using System;
using System.Data;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using MySql.Data.MySqlClient;

namespace controleDuvidas
{
    public partial class form_StatusResposta : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private string ocorrencia;
        private string analista;
        private bool ExisteImpacto;

        public form_StatusResposta(string codOcorrencia, string nomeAnalista, bool existeImpacto)
        {
            InitializeComponent();
            this.ocorrencia = codOcorrencia;
            this.analista = nomeAnalista;
            this.ExisteImpacto = existeImpacto;
        }

        private void form_StatusResposta_Load(object sender, EventArgs e)
        {
            try
            {
                Form1 f1 = new Form1();
                string equipe = ((f1.Equipe) == "Fábrica Desenvolvimento" ? "FD" : "FF");                
                
                //ABRE CONEXÃO
                bdConn.Open();
                
                //EXECUTA COMANDO
                MySqlCommand command = new MySqlCommand("SELECT * FROM status_resposta WHERE equipe like '%" + equipe + "%';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                    cb_StatusResposta.Items.Add(dr["nome_status"].ToString());                                               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
            finally
            {
                //FECHA CONEXÃO
                bdConn.Close(); 
            }
        }

        private void bt_Cancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bt_Atualizar_Click(object sender, EventArgs e)
        {
            if (cb_StatusResposta.Text != "")
            {
                try
                {
                    bool enviaEmail = false;
                    string statusAutomatico = "";
                    DialogResult result = DialogResult.Yes;
                    MySqlCommand command;
                    MySqlDataReader dr;

                    //ABRE CONEXÃO
                    bdConn.Open();

                    //VERIFICA EMAIL
                    command = new MySqlCommand("SELECT envia_email, status_automatico_sg FROM status_resposta WHERE nome_status = '" + cb_StatusResposta.Text + "';", bdConn);
                    dr = command.ExecuteReader();
                    if (dr.Read())
                    {
                        enviaEmail = ((dr["envia_email"].ToString() == "sim") ? true : false);
                        statusAutomatico = dr["status_automatico_sg"].ToString();
                    }
                    dr.Close();


                    if (ExisteImpacto == true)
                    {
                        //ATUALIAZA STATUS
                        command = new MySqlCommand("UPDATE ocorrencia SET status_resposta = '" + cb_StatusResposta.Text + "' WHERE cod_oco = '" + ocorrencia + "';", bdConn);
                        command.ExecuteNonQuery();

                        //ENVIA EMAIL
                        if (enviaEmail)
                        {
                            command = new MySqlCommand("SELECT O.identificador, O.status_resposta, O.status_geral, U.user_email FROM ocorrencia O INNER JOIN usuarios U ON O.identificador = U.user_nome AND O.cod_oco = '" + ocorrencia + "';", bdConn);
                            dr = command.ExecuteReader();
                            if (dr.Read())
                            {
                                //CRIA OBJETO DO EMAIL
                                Outlook.Application oApp = new Outlook.Application();
                                SendEmailFromAccount(oApp,
                                    "Alerta de Status - Controle de Ocorrências",
                                    criaEmailBody(dr["identificador"].ToString(), ocorrencia, dr["status_resposta"].ToString(), dr["status_geral"].ToString()),
                                    dr["user_email"].ToString(),
                                    (Environment.UserName.ToString() + "@accenture.com"));
                            }
                            dr.Close();
                        }

                        //CRIA REGISTOR DE LOG
                        command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + ocorrencia + "', '" + analista + "', 'Atualizacao Status', 'atualizou o Status Resposta para " + cb_StatusResposta.Text + " (SG: " + statusAutomatico + ")');", bdConn);
                        command.ExecuteNonQuery();

                        MessageBox.Show("Status Resposta foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    else
                    {

                        if (statusAutomatico != "")
                        {
                            result = MessageBox.Show("Ao prosseguir com o processo, o status geral será atualizado!\n\nStatus Geral: " + statusAutomatico + "\n\nDeseja continuar?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (result == DialogResult.Yes)
                            {
                                //ATUALIAZA STATUS
                                command = new MySqlCommand("UPDATE ocorrencia SET status_resposta = '" + cb_StatusResposta.Text + "' WHERE cod_oco = '" + ocorrencia + "';", bdConn);
                                command.ExecuteNonQuery();

                                //ATUALIAZA STATUS
                                command = new MySqlCommand("UPDATE ocorrencia SET status_geral = '" + statusAutomatico + "' WHERE cod_oco = '" + ocorrencia + "';", bdConn);
                                command.ExecuteNonQuery();

                                if (enviaEmail)
                                {
                                    command = new MySqlCommand("SELECT O.identificador, O.status_resposta, O.status_geral, U.user_email FROM ocorrencia O INNER JOIN usuarios U ON O.identificador = U.user_nome AND O.cod_oco = '" + ocorrencia + "';", bdConn);
                                    dr = command.ExecuteReader();
                                    if (dr.Read())
                                    {
                                        //CRIA OBJETO DO EMAIL
                                        Outlook.Application oApp = new Outlook.Application();
                                        SendEmailFromAccount(oApp,
                                            "Alerta de Status - Controle de Ocorrências",
                                            criaEmailBody(dr["identificador"].ToString(), ocorrencia, dr["status_resposta"].ToString(), dr["status_geral"].ToString()),
                                            dr["user_email"].ToString(),
                                            (Environment.UserName.ToString() + "@accenture.com"));
                                    }
                                    dr.Close();
                                }

                                //CRIA REGISTOR DE LOG
                                command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + ocorrencia + "', '" + analista + "', 'Atualizacao Status', 'atualizou o Status Resposta para " + cb_StatusResposta.Text + " (SG: " + statusAutomatico + ")');", bdConn);
                                command.ExecuteNonQuery();

                                MessageBox.Show("Status Resposta foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //ATUALIAZA STATUS
                            command = new MySqlCommand("UPDATE ocorrencia SET status_resposta = '" + cb_StatusResposta.Text + "' WHERE cod_oco = '" + ocorrencia + "';", bdConn);
                            command.ExecuteNonQuery();

                            if (enviaEmail)
                            {
                                command = new MySqlCommand("SELECT O.identificador, O.status_resposta, O.status_geral, U.user_email FROM ocorrencia O INNER JOIN usuarios U ON O.identificador = U.user_nome AND O.cod_oco = '" + ocorrencia + "';", bdConn);
                                dr = command.ExecuteReader();
                                if (dr.Read())
                                {
                                    //CRIA OBJETO DO EMAIL
                                    Outlook.Application oApp = new Outlook.Application();
                                    SendEmailFromAccount(oApp,
                                        "Alerta de Status - Controle de Ocorrências",
                                        criaEmailBody(dr["identificador"].ToString(), ocorrencia, dr["status_resposta"].ToString(), dr["status_geral"].ToString()),
                                        dr["user_email"].ToString(),
                                        (Environment.UserName.ToString() + "@accenture.com"));
                                }
                                dr.Close();
                            }

                            //CRIA REGISTOR DE LOG
                            command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + ocorrencia + "', '" + analista + "', 'Atualizacao Status', 'atualizou o Status Resposta para " + cb_StatusResposta.Text + " (SG: " + statusAutomatico + ")');", bdConn);
                            command.ExecuteNonQuery();

                            MessageBox.Show("Status Resposta foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (bdConn.State == ConnectionState.Open)
                        bdConn.Close();

                    this.Close();
                }
            }
            else
                MessageBox.Show("Selecione um Status para atualizar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //EMAIL - CRIA EMAIL
        static void SendEmailFromAccount(Outlook.Application application, string subject, string body, string to, string smtpAddress)
        {
            //CRIA ITEM DE EMAIL: DE, PARA, ASSUNTO, CORPO       
            Outlook.MailItem newMail = (Outlook.MailItem)application.CreateItem(Outlook.OlItemType.olMailItem);


            newMail.To = to;                                                            //PARA            
            newMail.Subject = subject;                                                  //ASSUNTO
            newMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;                     //--------
            newMail.HTMLBody = body;                                                    //CORPO DO EMAIL
            newMail.Importance = Outlook.OlImportance.olImportanceHigh;                 //IMPORTANCIA

            // Recuperar a conta que tem o endereço SMTP específico.
            Outlook.Account account = GetAccountForEmailAddress(application, smtpAddress);
            // Usa conta para enviar o e- mail.
            newMail.SendUsingAccount = account;
            newMail.Send();
        }

        //EMAIL - VERIFICA CONTAS OUTLOOK
        static Outlook.Account GetAccountForEmailAddress(Outlook.Application application, string smtpAddress)
        {

            //Loop sobre a coleção de Contas da sessão atual do Outlook.
            Outlook.Accounts accounts = application.Session.Accounts;
            foreach (Outlook.Account account in accounts)
            {
                // When the e-mail address matches, return the account.
                if (account.SmtpAddress == smtpAddress)
                    return account;
            }
            throw new System.Exception(string.Format("Email : {0} registrado no 'Envio de Email' não existe ou está incorreto! Favor desativar e ativar o 'Envio de Email' novamente.", smtpAddress));
        }

        //EMAIL - CRIA CORPO DO EMAIL
        static string criaEmailBody(string responsavel, string ocorrencia, string statusResposta, string statusGeral)
        {
            string emailBody = "";

            #region COMENTARIOS
            emailBody += "<html xmlns:o='urn:schemas-microsoft-com:office:office'";
            emailBody += "xmlns:w='urn:schemas-microsoft-com:office:word'";
            emailBody += "xmlns:m='http://schemas.microsoft.com/office/2004/12/omml'";
            emailBody += "xmlns='http://www.w3.org/TR/REC-html40'>";
            emailBody += "";
            emailBody += "<head>";
            emailBody += "<meta http-equiv=Content-Type content='text/html; charset=windows-1252'>";
            emailBody += "<meta name=ProgId content=Word.Document>";
            emailBody += "<meta name=Generator content='Microsoft Word 15'>";
            emailBody += "<meta name=Originator content='Microsoft Word 15'>";
            emailBody += "<link rel=File-List href='Sem%20título_arquivos/filelist.xml'>";
            emailBody += "<link rel=Edit-Time-Data href='Sem%20título_arquivos/editdata.mso'>";
            emailBody += "<link rel=themeData href='Sem%20título_arquivos/themedata.thmx'>";
            emailBody += "<link rel=colorSchemeMapping href='Sem%20título_arquivos/colorschememapping.xml'>";
            emailBody += "<!--[if gte mso 9]><xml>";
            emailBody += " <w:WordDocument>";
            emailBody += "  <w:Zoom>0</w:Zoom>";
            emailBody += "  <w:TrackMoves/>";
            emailBody += "  <w:TrackFormatting/>";
            emailBody += "  <w:HyphenationZone>21</w:HyphenationZone>";
            emailBody += "  <w:ValidateAgainstSchemas/>";
            emailBody += "  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>";
            emailBody += "  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>";
            emailBody += "  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>";
            emailBody += "  <w:DoNotPromoteQF/>";
            emailBody += "  <w:LidThemeOther>PT-BR</w:LidThemeOther>";
            emailBody += "  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>";
            emailBody += "  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>";
            emailBody += "  <w:Compatibility>";
            emailBody += "   <w:DoNotExpandShiftReturn/>";
            emailBody += "   <w:BreakWrappedTables/>";
            emailBody += "   <w:SplitPgBreakAndParaMark/>";
            emailBody += "   <w:EnableOpenTypeKerning/>";
            emailBody += "  </w:Compatibility>";
            emailBody += "  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>";
            emailBody += "  <m:mathPr>";
            emailBody += "   <m:mathFont m:val='Cambria Math'/>";
            emailBody += "   <m:brkBin m:val='before'/>";
            emailBody += "   <m:brkBinSub m:val='&#45;-'/>";
            emailBody += "   <m:smallFrac m:val='off'/>";
            emailBody += "   <m:dispDef/>";
            emailBody += "   <m:lMargin m:val='0'/>";
            emailBody += "   <m:rMargin m:val='0'/>";
            emailBody += "   <m:defJc m:val='centerGroup'/>";
            emailBody += "   <m:wrapIndent m:val='1440'/>";
            emailBody += "   <m:intLim m:val='subSup'/>";
            emailBody += "   <m:naryLim m:val='undOvr'/>";
            emailBody += "  </m:mathPr></w:WordDocument>";
            emailBody += "</xml><![endif]--><!--[if gte mso 9]><xml>";
            emailBody += " <w:LatentStyles DefLockedState='false' DefUnhideWhenUsed='false'";
            emailBody += "  DefSemiHidden='false' DefQFormat='false' DefPriority='99'";
            emailBody += "  LatentStyleCount='371'>";
            emailBody += "  <w:LsdException Locked='false' Priority='0' QFormat='true' Name='Normal'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' QFormat='true' Name='heading 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 7'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 8'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='9' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='heading 9'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 6'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 7'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 8'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index 9'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 7'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 8'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='toc 9'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Normal Indent'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='footnote text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='annotation text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='header'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='footer'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='index heading'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='35' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='caption'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='table of figures'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='envelope address'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='envelope return'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='footnote reference'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='annotation reference'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='line number'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='page number'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='endnote reference'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='endnote text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='table of authorities'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='macro'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='toa heading'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Bullet'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Number'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Bullet 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Bullet 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Bullet 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Bullet 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Number 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Number 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Number 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Number 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='10' QFormat='true' Name='Title'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Closing'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Signature'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='1' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='Default Paragraph Font'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text Indent'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Continue'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Continue 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Continue 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Continue 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='List Continue 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Message Header'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='11' QFormat='true' Name='Subtitle'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Salutation'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Date'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text First Indent'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text First Indent 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Note Heading'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text Indent 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Body Text Indent 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Block Text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Hyperlink'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='FollowedHyperlink'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='22' QFormat='true' Name='Strong'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='20' QFormat='true' Name='Emphasis'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Document Map'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Plain Text'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='E-mail Signature'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Top of Form'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Bottom of Form'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Normal (Web)'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Acronym'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Address'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Cite'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Code'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Definition'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Keyboard'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Preformatted'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Sample'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Typewriter'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='HTML Variable'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Normal Table'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='annotation subject'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='No List'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Outline List 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Outline List 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Outline List 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Simple 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Simple 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Simple 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Classic 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Classic 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Classic 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Classic 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Colorful 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Colorful 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Colorful 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Columns 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Columns 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Columns 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Columns 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Columns 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 6'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 7'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Grid 8'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 4'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 5'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 6'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 7'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table List 8'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table 3D effects 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table 3D effects 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table 3D effects 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Contemporary'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Elegant'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Professional'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Subtle 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Subtle 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Web 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Web 2'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Web 3'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Balloon Text'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' Name='Table Grid'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' UnhideWhenUsed='true'";
            emailBody += "   Name='Table Theme'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' Name='Placeholder Text'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='1' QFormat='true' Name='No Spacing'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' SemiHidden='true' Name='Revision'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='34' QFormat='true'";
            emailBody += "   Name='List Paragraph'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='29' QFormat='true' Name='Quote'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='30' QFormat='true'";
            emailBody += "   Name='Intense Quote'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='60' Name='Light Shading Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='61' Name='Light List Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='62' Name='Light Grid Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='63' Name='Medium Shading 1 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='64' Name='Medium Shading 2 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='65' Name='Medium List 1 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='66' Name='Medium List 2 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='67' Name='Medium Grid 1 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='68' Name='Medium Grid 2 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='69' Name='Medium Grid 3 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='70' Name='Dark List Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='71' Name='Colorful Shading Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='72' Name='Colorful List Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='73' Name='Colorful Grid Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='19' QFormat='true'";
            emailBody += "   Name='Subtle Emphasis'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='21' QFormat='true'";
            emailBody += "   Name='Intense Emphasis'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='31' QFormat='true'";
            emailBody += "   Name='Subtle Reference'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='32' QFormat='true'";
            emailBody += "   Name='Intense Reference'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='33' QFormat='true' Name='Book Title'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='37' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' Name='Bibliography'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='39' SemiHidden='true'";
            emailBody += "   UnhideWhenUsed='true' QFormat='true' Name='TOC Heading'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='41' Name='Plain Table 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='42' Name='Plain Table 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='43' Name='Plain Table 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='44' Name='Plain Table 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='45' Name='Plain Table 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='40' Name='Grid Table Light'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46' Name='Grid Table 1 Light'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51' Name='Grid Table 6 Colorful'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52' Name='Grid Table 7 Colorful'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='Grid Table 1 Light Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='Grid Table 2 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='Grid Table 3 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='Grid Table 4 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='Grid Table 5 Dark Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='Grid Table 6 Colorful Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='Grid Table 7 Colorful Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46' Name='List Table 1 Light'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51' Name='List Table 6 Colorful'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52' Name='List Table 7 Colorful'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 1'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 2'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 3'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 4'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 5'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='46'";
            emailBody += "   Name='List Table 1 Light Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='47' Name='List Table 2 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='48' Name='List Table 3 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='49' Name='List Table 4 Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='50' Name='List Table 5 Dark Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='51'";
            emailBody += "   Name='List Table 6 Colorful Accent 6'/>";
            emailBody += "  <w:LsdException Locked='false' Priority='52'";
            emailBody += "   Name='List Table 7 Colorful Accent 6'/>";
            emailBody += " </w:LatentStyles>";
            emailBody += "</xml><![endif]-->";
            emailBody += "<style>";
            emailBody += "<!--";
            emailBody += " /* Font Definitions */";
            emailBody += " @font-face";
            emailBody += "	{font-family:'Cambria Math';";
            emailBody += "	panose-1:2 4 5 3 5 4 6 3 2 4;";
            emailBody += "	mso-font-charset:1;";
            emailBody += "	mso-generic-font-family:roman;";
            emailBody += "	mso-font-format:other;";
            emailBody += "	mso-font-pitch:variable;";
            emailBody += "	mso-font-signature:0 0 0 0 0 0;}";
            emailBody += "@font-face";
            emailBody += "	{font-family:Calibri;";
            emailBody += "	panose-1:2 15 5 2 2 2 4 3 2 4;";
            emailBody += "	mso-font-charset:0;";
            emailBody += "	mso-generic-font-family:swiss;";
            emailBody += "	mso-font-pitch:variable;";
            emailBody += "	mso-font-signature:-536870145 1073786111 1 0 415 0;}";
            emailBody += "@font-face";
            emailBody += "	{font-family:'Arial Rounded MT Bold';";
            emailBody += "	panose-1:2 15 7 4 3 5 4 3 2 4;";
            emailBody += "	mso-font-charset:0;";
            emailBody += "	mso-generic-font-family:swiss;";
            emailBody += "	mso-font-pitch:variable;";
            emailBody += "	mso-font-signature:3 0 0 0 1 0;}";
            emailBody += " /* Style Definitions */";
            emailBody += " p.MsoNormal, li.MsoNormal, div.MsoNormal";
            emailBody += "	{mso-style-unhide:no;";
            emailBody += "	mso-style-qformat:yes;";
            emailBody += "	mso-style-parent:'';";
            emailBody += "	margin:0cm;";
            emailBody += "	margin-bottom:.0001pt;";
            emailBody += "	mso-pagination:widow-orphan;";
            emailBody += "	font-size:11.0pt;";
            emailBody += "	font-family:'Calibri',sans-serif;";
            emailBody += "	mso-fareast-font-family:Calibri;";
            emailBody += "	mso-fareast-theme-font:minor-latin;";
            emailBody += "	mso-bidi-font-family:'Times New Roman';";
            emailBody += "	mso-fareast-language:EN-US;}";
            emailBody += "a:link, span.MsoHyperlink";
            emailBody += "	{mso-style-noshow:yes;";
            emailBody += "	mso-style-priority:99;";
            emailBody += "	color:#0563C1;";
            emailBody += "	text-decoration:underline;";
            emailBody += "	text-underline:single;}";
            emailBody += "a:visited, span.MsoHyperlinkFollowed";
            emailBody += "	{mso-style-noshow:yes;";
            emailBody += "	mso-style-priority:99;";
            emailBody += "	color:#954F72;";
            emailBody += "	text-decoration:underline;";
            emailBody += "	text-underline:single;}";
            emailBody += "span.EstiloDeEmail17";
            emailBody += "	{mso-style-type:personal;";
            emailBody += "	mso-style-noshow:yes;";
            emailBody += "	mso-style-unhide:no;";
            emailBody += "	font-family:'Calibri',sans-serif;";
            emailBody += "	mso-ascii-font-family:Calibri;";
            emailBody += "	mso-hansi-font-family:Calibri;";
            emailBody += "	color:#1F4E79;}";
            emailBody += ".MsoChpDefault";
            emailBody += "	{mso-style-type:export-only;";
            emailBody += "	mso-default-props:yes;";
            emailBody += "	font-size:10.0pt;";
            emailBody += "	mso-ansi-font-size:10.0pt;";
            emailBody += "	mso-bidi-font-size:10.0pt;}";
            emailBody += "@page WordSection1";
            emailBody += "	{size:612.0pt 792.0pt;";
            emailBody += "	margin:70.85pt 3.0cm 70.85pt 3.0cm;";
            emailBody += "	mso-header-margin:36.0pt;";
            emailBody += "	mso-footer-margin:36.0pt;";
            emailBody += "	mso-paper-source:0;}";
            emailBody += "div.WordSection1";
            emailBody += "	{page:WordSection1;}";
            emailBody += "-->";
            emailBody += "</style>";
            emailBody += "<!--[if gte mso 10]>";
            emailBody += "<style>";
            emailBody += " /* Style Definitions */";
            emailBody += " table.MsoNormalTable";
            emailBody += "	{mso-style-name:'Tabela normal';";
            emailBody += "	mso-tstyle-rowband-size:0;";
            emailBody += "	mso-tstyle-colband-size:0;";
            emailBody += "	mso-style-noshow:yes;";
            emailBody += "	mso-style-priority:99;";
            emailBody += "	mso-style-parent:'';";
            emailBody += "	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;";
            emailBody += "	mso-para-margin:0cm;";
            emailBody += "	mso-para-margin-bottom:.0001pt;";
            emailBody += "	mso-pagination:widow-orphan;";
            emailBody += "	font-size:10.0pt;";
            emailBody += "	font-family:'Times New Roman',serif;}";
            emailBody += "</style>";
            emailBody += "<![endif]-->";
            emailBody += "</head>";
            emailBody += "";
            #endregion

            emailBody += "<body lang=PT-BR link='#0563C1' vlink='#954F72' style='tab-interval:35.4pt'>";
            emailBody += "";
            emailBody += "<div class=WordSection1>";
            emailBody += "";
            emailBody += "<p class=MsoNormal><o:p>&nbsp;</o:p></p>";
            emailBody += "";
            emailBody += "<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0";
            emailBody += " style='border-collapse:collapse;mso-yfti-tbllook:1184;mso-padding-alt:0cm 0cm 0cm 0cm'>";
            emailBody += " <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:45.0pt'>";
            emailBody += "  <td width=756 style='width:20.0cm;background:#171717;padding:0cm 5.4pt 0cm 5.4pt; height:45.0pt'>  ";
            emailBody += "  <p class=MsoNormal align=center style='text-align:center'><span style='font-size:30.0pt'>Controle de Ocorrências<o:p></o:p></span></p>";
            emailBody += "  </td>";
            emailBody += " </tr>";
            emailBody += " <tr style='mso-yfti-irow:1;mso-yfti-lastrow:yes;height:241.75pt'>";
            emailBody += "  <td width=756 valign=top style='width:20.0cm;background:#EDEDED;padding:0cm 5.4pt 0cm 5.4pt; height:241.75pt'>  ";
            emailBody += "  <p class=MsoNormal><o:p>&nbsp;</o:p></p>";
            emailBody += "  <p class=MsoNormal><o:p>&nbsp;</o:p></p>";
            emailBody += "  ";
            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Olá " + responsavel + ",<o:p></o:p></span></p>  ";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Sua ocorrência foi atualizada.<o:p></o:p></span></p>  ";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>Código Ocorrência: </b><span style='color:red'>" + ocorrencia + "</span><o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>Status Resposta: </b><span style='color:black'>" + statusResposta + "</span><o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>Status Geral: </b><span style='color:black'>" + statusGeral + "</span><o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Favor acessar a ferramenta “<i>Controle de Ocorrências</i>” para outros detalhes.<o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Lembre-se de verificar se a resposta foi satisfatória e o problema já foi solucionado, caso sim, atualize o “<i>Status Geral</i>” para finalizar a ocorrência.<o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Qualquer dúvidas solicitar apoio da gestão.<o:p></o:p></span></p>";
            emailBody += "  ";
            emailBody += "  </td>";
            emailBody += " </tr>";
            emailBody += "</table>";
            emailBody += "";
            emailBody += "<p class=MsoNormal><o:p>&nbsp;</o:p></p>";
            emailBody += "";
            emailBody += "<p class=MsoNormal><i>Esta é uma mensagem automática. Solicitamos, por favor, não responder este e-mail.</i></p>";
            emailBody += "";
            emailBody += "<p class=MsoNormal style='background:white'><span style='font-size:12.0pt; font-family:'Times New Roman',serif;color:#222222;mso-fareast-language:PT-BR'><o:p>&nbsp;</o:p></span></p>";
            emailBody += "";
            emailBody += "<p class=MsoNormal><o:p>&nbsp;</o:p></p>";
            emailBody += "";
            emailBody += "</div>";
            emailBody += "";
            emailBody += "</body>";
            emailBody += "";
            emailBody += "</html>";


            return emailBody;
        }
    
    }
}
