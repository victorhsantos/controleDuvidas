using System;
using System.Data;
using System.IO;
using MySql.Data.MySqlClient;
using System.Reflection;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using controleOcorrencias;

namespace controleDuvidas
{
    public partial class Form1 : Form
    {
        //BANCO DE DADOS
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private MySqlDataAdapter bdAdapter;
        private DataSet bdDataSet;
        ToolTip buttonToolTip = new ToolTip();
        static FormWindowState stateAtual = FormWindowState.Normal;
        static string ocoPainel = "";
        static bool userLogado = false;
        string configVoltar = "";

        //ControladoresMO Anexos
        public static int countAnexo = 0;
        public static int remAnexo = 0;

        //LinkLabel Anexo Registro Ocorrência
        private string caminhoOrigemA1;
        private string caminhoOrigemA2;
        private string caminhoOrigemA3;
        private string caminhoOrigemA4;
        private string caminhoOrigemA5;

        //LinkLabel Anexo Detalhamento
        private string caminhoOrigemA1D;
        private string caminhoOrigemA2D;
        private string caminhoOrigemA3D;
        private string caminhoOrigemA4D;
        private string caminhoOrigemA5D;

        private static string UserON_Equipe = "";
        private static string UserON_TipoUser = "";
        private static bool controle_ComboboxDO = false;

        public int removeAnexo { set { remAnexo = value; } }

        public string Equipe { get { return UserON_Equipe; } }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //FORM PRINCIPAL - RECUPERANDO A VERSÃO
            Version versao = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = "Controle de Ocorrências - Versão " + versao.ToString().Substring(0, 3);

            panel_Inicio.Dock = DockStyle.Fill;
            panel_MenuParametros.Dock = DockStyle.Fill;
            panel_CadastroProjeto.Dock = DockStyle.Fill;
            panel_MO_MinhasOcorrencias.Dock = DockStyle.Fill;
            panel_OcorrenciaDetalhes.Dock = DockStyle.Fill;
            panel_MO_BuscarProjetos1.Dock = DockStyle.Fill;
            panel_MO_OID.Dock = DockStyle.Fill;

            tb_User.Text = Environment.UserName;

            buttonToolTip.SetToolTip(pictureBox1, "Usuário");
            buttonToolTip.SetToolTip(pictureBox2, "Senha");

            string versaoAt = "";
            string linkAt = "";

            try
            {
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT versao, link_atualizacao FROM versao_co ORDER BY versao DESC LIMIT 1;", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                if (dr.Read())
                {
                    versaoAt = dr["versao"].ToString();
                    linkAt = dr["link_atualizacao"].ToString();
                }
                dr.Close();
            }
            catch
            {
                MessageBox.Show("Erro ao verificar a versão da ferramenta.", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                bdConn.Close();
            }

            if (versaoAt != "" && linkAt != "")
                if (Double.Parse(versaoAt) > Double.Parse(versao.ToString().Substring(0, 3)))
                {
                    DialogResult result = MessageBox.Show("Existe uma nova versão da ferramenta disponível.\n\nDeseja baixar agora?", "Atualização Disponível!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("explorer.exe", linkAt);
                        this.Close();
                    }
                }
        }

        #region //****************************************** LOGIN  ******************************************\\

        //BOTÃO PARA CADASTRO DE USUARIO
        private void lb_CadastrarUser_Click(object sender, EventArgs e)
        {
            FormCadastroUser cadastro = new FormCadastroUser();
            cadastro.ShowDialog();
        }

        //BOTÃO PARA RECUPERAR SENHA
        private void lb_EsqueciSenha_Click(object sender, EventArgs e)
        {
            form_RecuperarSenha fRecuperarSenha = new form_RecuperarSenha();
            fRecuperarSenha.ShowDialog();
        }

        //BOTÃO PARA ENTRAR
        private void bt_SingIn_Click(object sender, EventArgs e)
        {
            try
            {
                bool userOK = false;
                bool viewRelatorio = false;

                //ABRE CONEXÃO
                bdConn.Open();

                MySqlCommand comand = new MySqlCommand("SELECT user_senha, user_nome, user_lvl, user_equipe, permissao_relatorio, view_relatorio FROM usuarios WHERE user_login = '" + tb_User.Text + "';", bdConn);
                MySqlDataReader dr = comand.ExecuteReader();

                if (dr.Read())
                    if (tb_Password.Text == dr["user_senha"].ToString())
                    {
                        UserON_Equipe = dr["user_equipe"].ToString();
                        UserON_TipoUser = ((dr["user_lvl"].ToString() == "0") ? "Comum" : "Administrador");
                        lb_UserON.Text = dr["user_nome"].ToString();
                        configuraçõesToolStripMenuItem.Enabled = ((dr["user_lvl"].ToString() == "0") ? false : true);
                        relatórioDeOcorrênciasToolStripMenuItem.Enabled = ((dr["permissao_relatorio"].ToString() == "nao") ? false : true);
                        viewRelatorio = ((dr["view_relatorio"].ToString() == "nao") ? false : true);
                        userOK = true;
                    }
                    else
                        throw new Exception("Senha incorreta!");
                else
                    throw new Exception("Usuário não existe ou está incorreto!");

                //FECHA CONEXÃO
                bdConn.Close();

                if (viewRelatorio)
                {
                    form_RelatorioGestao frRelatorio = new form_RelatorioGestao(lb_UserON.Text, "Por Status");
                    frRelatorio.ShowDialog();
                }

                if (userOK)
                {
                    home();

                    buttonToolTip.UseFading = true;
                    buttonToolTip.UseAnimation = true;
                    buttonToolTip.IsBalloon = true;

                    buttonToolTip.ShowAlways = true;

                    buttonToolTip.AutoPopDelay = 5000;
                    buttonToolTip.InitialDelay = 1000;
                    buttonToolTip.ReshowDelay = 500;

                    buttonToolTip.SetToolTip(pictureBox_UserOn, "Usuário Online");
                    buttonToolTip.SetToolTip(bt_AtualizarPainelOcorrencias, "Atualizar Painel de Ocorrências");

                    userLogado = true;
                    tb_User.Focus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                tb_Password.Text = "";
                bdConn.Close();
            }
        }

        //VERIFICA ENTER
        private void tb_User_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
                bt_SingIn.PerformClick();
        }

        //VERIFICA ENTER
        private void tb_Password_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
                bt_SingIn.PerformClick();
        }

        #endregion

        #region //****************************************** HOME  ******************************************\\

        //FUNÇÃO INICIO
        void home()
        {
            panel_Inicio.BringToFront();
            panel_Inicio.Visible = true;
            panel_RegOcorrencia_1.Visible = false;
            panel_MenuParametros.Visible = false;
            panel_CadastroProjeto.Visible = false;
            panel_MO_MinhasOcorrencias.Visible = false;
            panel_OcorrenciaDetalhes.Visible = false;
            panel_MO_BuscarProjetos1.Visible = false;
            panel_MO_OID.Visible = false;

            atualizaPainelOcorrencias();
        }

        //ATUALIZA PAINEL DE OCORRÊNCIAS
        void atualizaPainelOcorrencias()
        {
            try
            {
                //LIMPA GRIDVIEW
                if (this.dataGrid_PainelOcorrencias.DataSource != null)
                    this.dataGrid_PainelOcorrencias.DataSource = null;
                else
                {
                    this.dataGrid_PainelOcorrencias.Rows.Clear();
                    this.dataGrid_PainelOcorrencias.Columns.Clear();
                }

                //HEADER DATAGRIDVIEW
                dataGrid_PainelOcorrencias.Columns.Add("cod_oco", "PRJ - Ocorrência");
                dataGrid_PainelOcorrencias.Columns.Add("sistema", "Sistema");
                dataGrid_PainelOcorrencias.Columns.Add("dfrqf", "RQF");
                dataGrid_PainelOcorrencias.Columns.Add("identificador", "Identificador");
                dataGrid_PainelOcorrencias.Columns.Add("status_geral", "Status Geral");
                dataGrid_PainelOcorrencias.Columns.Add("atrib_para", "Atribuído Para");
                dataGrid_PainelOcorrencias.Columns.Add("impacto_cttu", "Impacto CTTU");
                dataGrid_PainelOcorrencias.Columns.Add("data_result", "Ult. Atualização");


                //ABRE CONEXÃO
                bdConn.Open();
                MySqlCommand comand;

                if (this.WindowState == FormWindowState.Maximized)
                    comand = new MySqlCommand("SELECT A.cod_oco, A.sistema, A.cod_rqf, A.identificador, A.atrib_para, A.status_geral, A.status_resposta, A.impacto_cttu, IF( A.horario_registro < B.ultima_data, B.ultima_data, A.horario_registro ) AS DATA_RESULT FROM (SELECT * FROM ocorrencia) AS A JOIN (SELECT cod_oco, max( horario_registro ) AS ultima_data FROM desc_ocorrencia GROUP BY cod_oco ORDER BY ultima_data) AS B ON A.cod_oco = B.cod_oco ORDER BY DATA_RESULT DESC LIMIT 28;", bdConn);
                else
                    comand = new MySqlCommand("SELECT A.cod_oco, A.sistema, A.cod_rqf, A.identificador, A.atrib_para, A.status_geral, A.status_resposta, A.impacto_cttu, IF( A.horario_registro < B.ultima_data, B.ultima_data, A.horario_registro ) AS DATA_RESULT FROM (SELECT * FROM ocorrencia) AS A JOIN (SELECT cod_oco, max( horario_registro ) AS ultima_data FROM desc_ocorrencia GROUP BY cod_oco ORDER BY ultima_data) AS B ON A.cod_oco = B.cod_oco ORDER BY DATA_RESULT DESC LIMIT 23;", bdConn);

                MySqlDataReader dr = comand.ExecuteReader();

                while (dr.Read())
                {
                    DateTime ultAtul = DateTime.Parse(dr["DATA_RESULT"].ToString());
                    dataGrid_PainelOcorrencias.Rows.Add(
                        dr["cod_oco"].ToString(),
                        dr["sistema"].ToString(),
                        dr["cod_rqf"].ToString(),
                        dr["identificador"].ToString(),
                        dr["status_geral"].ToString(),
                        dr["atrib_para"].ToString(),
                        dr["impacto_cttu"].ToString(),
                        String.Format("{0:dd/MM/yyyy}", ultAtul)
                        );
                    //dr["DATA_RESULT"].ToString().Substring(0, 5) + " - " + dr["DATA_RESULT"].ToString().Substring(11, 5)
                }

                dataGrid_PainelOcorrencias.ClearSelection();

                //FECHA CONEXÃO
                bdConn.Close();

                //LABEL ATUALIZADO
                lb_PainelOcorrencia_Atualizacao.Text = "Painel Atualizado em: " + DateTime.Now;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                bdConn.Close();
            }
        }

        //VERIFICA SE ESTÁ MAXI OU NORMAL
        private void Form1_Resize(object sender, EventArgs e)
        {
            if (userLogado)
                if ((this.WindowState != stateAtual) && (this.WindowState != FormWindowState.Minimized))
                {
                    stateAtual = this.WindowState;
                    atualizaPainelOcorrencias();
                }
        }

        //PAINEL DE CONTROLES
        private void painelDeOcorrênciasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            home();
        }

        //MENU CONFIGURAÇÕES
        private void configuraçõesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //REGISTRO DE OCORRENCIA       
        private void novaOcorrênciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            inicioRegistroOcorrencia();
        }

        //MINHAS OCORRENCIAS
        private void minhasOcorrêciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel_MO_MinhasOcorrencias.BringToFront();

            panel_MO_MinhasOcorrencias.Visible = true;

            inicioMinhasOcorrencias();
        }

        //RELATÓRIO OCORRÊNCIAS - POR STATUS
        private void porStatusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form_RelatorioGestao frRelatorio = new form_RelatorioGestao(lb_UserON.Text, "Por Status");
            frRelatorio.ShowDialog();
        }

        //RELATÓRIO OCORRÊNCIAS - POR OCORRÊNCIA
        private void porOcorrênciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form_RelatorioGestao frRelatorio = new form_RelatorioGestao(lb_UserON.Text, "Por Ocorrência");
            frRelatorio.ShowDialog();
        }

        //PESQUISAR OCORRENCIAS POR ID
        private void porIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel_MO_OID.BringToFront();

            panel_MO_OID.Visible = true;

            tb_OID_SelectID.Text = "P_____-OC___";
            tb_OID_SelectID.ForeColor = Color.Gray;
        }

        //PESQUISAR OCORRENCIAS POR PROJETO
        private void porProjetoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel_MO_BuscarProjetos1.BringToFront();

            panel_MO_BuscarProjetos1.Visible = true;
            panel_MO_BuscarProjetos2.Visible = false;

            tb_OP_SelectPRJ.Text = "";
        }

        //BOTÃO LOGOUT
        private void desconectarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            panel_Inicio.Visible = false;
            lb_UserON.Text = "";

            tb_Password.Text = "";

            userLogado = false;
        }

        //BOTÃO ALTERAR SENHA
        private void alterarSenhaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form_AlterarSenha fAlterarSenha = new form_AlterarSenha(lb_UserON.Text);
            fAlterarSenha.ShowDialog();
        }

        //BOTÃO PARA ATUALIZA PAINEL DE OCORRÊNCIAS
        private void bt_AtualizarPainelOcorrencias_Click(object sender, EventArgs e)
        {
            atualizaPainelOcorrencias();
        }

        //FORMATA DATAGRIDVIEW
        private void dataGrid_PainelOcorrencias_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.Value != null) && (e.ColumnIndex == 3 || e.ColumnIndex == 5))
            {
                if (e.Value.Equals(lb_UserON.Text))
                    e.CellStyle.BackColor = Color.Gold;
            }
        }

        //CLICK DATAGRIDVIEW - BOTÃO ABRIR
        private void dataGrid_PainelOcorrencias_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGrid_PainelOcorrencias.CurrentRow != null)
            {
                inicioMenuOcorrencias();
                abreDetalhamentoOcorrencia(dataGrid_PainelOcorrencias.CurrentRow.Cells[0].Value.ToString(), "home");
            }
        }

        //INFO
        private void sobreToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            controleOcorrencias.formInfo frI = new controleOcorrencias.formInfo();
            frI.ShowDialog();
        }

        //MENU CLICK BOTÃO DIREITO PAINEL DE OCORRÊNCIAS
        private void dataGrid_PainelOcorrencias_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ContextMenu m = new ContextMenu();
                MenuItem MI_abreOco = new MenuItem("Abrir Ocorrência");
                m.MenuItems.Add(MI_abreOco);

                int currentMouseOverRow = dataGrid_PainelOcorrencias.HitTest(e.X, e.Y).RowIndex;
                if (currentMouseOverRow >= 0)
                {
                    //currentMouseOverRow.ToString();
                    m.Show(dataGrid_PainelOcorrencias, new Point(e.X, e.Y));
                    ocoPainel = dataGrid_PainelOcorrencias.CurrentRow.Cells[0].Value.ToString();
                    MI_abreOco.Click += new EventHandler(this.MI_abreOco_Click);
                }
            }
        }

        //MENU CLICK BOTÃO DIREITO - ABRE OCORRÊNCIA
        private void MI_abreOco_Click(object sender, System.EventArgs e)
        {
            if (ocoPainel != "")
                abreDetalhamentoOcorrencia(ocoPainel, "home");
            ocoPainel = "";
        }

        //AJUDA SLA
        private void pictureBox16_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Adicionar o tempo de SLA para os Impacto CTTU: Baixa, Média e Alta. Lembrando que este será o prazo para as respostas das ocorrências deste projeto!", "Tempo para SLA", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region //****************************************** PARAMETROS  ******************************************\\

        //INICIO - MENU PARAMETROS
        void inicioMenuParametros()
        {
            panel_MenuParametros.BringToFront();
            panel_CadastroProjeto.SendToBack();
            panel_EditarUser.SendToBack();
            panel_CadastrarRQF_1.SendToBack();
            panel_CadastrarRQF_2.SendToBack();
            panel_EditarRQF.SendToBack();
            panel_NovoStatus1.SendToBack();
            panel_NovoStatus2.SendToBack();

            panel_MenuParametros.Visible = true;
            panel_CadastroProjeto.Visible = false;
            panel_EditarUser.Visible = false;
            panel_CadastrarRQF_1.Visible = false;
            panel_CadastrarRQF_2.Visible = false;
            panel_EditarRQF.Visible = false;
            panel_NovoStatus1.Visible = false;
            panel_NovoStatus2.Visible = false;

            panel_MenuParametros.Dock = DockStyle.Fill;
            panel_CadastroProjeto.Dock = DockStyle.Fill;
            panel_EditarUser.Dock = DockStyle.Fill;
            panel_CadastrarRQF_1.Dock = DockStyle.Fill;
            panel_CadastrarRQF_2.Dock = DockStyle.Fill;
            panel_EditarRQF.Dock = DockStyle.Fill;
            panel_NovoStatus1.Dock = DockStyle.Fill;
            panel_NovoStatus2.Dock = DockStyle.Fill;
        }
        
        //BOTÃO RETURN - MENU PARAMETROS
        private void bt_Parametros_Return_Click(object sender, EventArgs e)
        {
            panel_MenuParametros.SendToBack();
            panel_CadastroProjeto.SendToBack();
            panel_EditarUser.SendToBack();
            panel_CadastrarRQF_1.SendToBack();
            panel_CadastrarRQF_2.SendToBack();
            panel_EditarRQF.SendToBack();
            panel_NovoStatus1.SendToBack();
            panel_NovoStatus2.SendToBack();

            panel_MenuParametros.Visible = false;
            panel_CadastroProjeto.Visible = false;
            panel_EditarUser.Visible = false;
            panel_CadastrarRQF_1.Visible = false;
            panel_CadastrarRQF_2.Visible = false;
            panel_EditarRQF.Visible = false;
            panel_NovoStatus1.Visible = false;
            panel_NovoStatus2.Visible = false;

            atualizaPainelOcorrencias();
        }

        //BOTÃO CADASTRR PROJETO - MENU PARAMETROS
        private void bt_CadastarProjeto_Click(object sender, EventArgs e)
        {
            //HABILITA PANEL
            panel_CadastroProjeto.Visible = true;
            panel_CadastroProjeto.BringToFront();

            //CONF. TEXTBOX
            tb_Projeto_CadastroPRJ.Visible = true;
            tb_Projeto_CadastroPRJ.BringToFront();
            cb_EditarPRJ_Projeto.Visible = false;

            //CONF BOTÕES
            bt_Salvar_CadastroPRJ.Visible = true;
            bt_EditarPRJ_Editar.Visible = false;

            lb_Titulo_CadastarProjeto.Visible = true;
            lb_Titulo_EditarProjeto.Visible = false;
            pb_EditarPRJ.Visible = false;
            pb_CadastrarPRJ.Visible = true;

            limpaCadastroPRJ();

            #region ATUALIZA COMBOBOX DE ANALISTAS
            if (bdConn.State == ConnectionState.Closed)
            {
                try
                {
                    //LIMPA COMBOBOX
                    cb_LTFabrica_CadastroPRJ.Items.Clear();
                    cb_LTFuncional_CadastroPRJ.Items.Clear();

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Desenvolvimento';", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_LTFabrica_CadastroPRJ.Items.Add(dr["user_nome"].ToString());
                    dr.Close();

                    command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Funcional';", bdConn);
                    dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_LTFuncional_CadastroPRJ.Items.Add(dr["user_nome"].ToString());
                    dr.Close();


                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            #endregion

        }

        //BOTÃO EDITAR PROJETO - MENU PARAMETROS
        private void bt_EditarProjeto_Click(object sender, EventArgs e)
        {
            //HABILITA PANEL                                
            panel_CadastroProjeto.Visible = true;
            panel_CadastroProjeto.BringToFront();

            //CONF. TEXTBOX
            cb_EditarPRJ_Projeto.Visible = true;
            cb_EditarPRJ_Projeto.BringToFront();
            tb_Projeto_CadastroPRJ.Visible = false;

            //CONF BOTÕES
            bt_Salvar_CadastroPRJ.Visible = false;
            bt_EditarPRJ_Editar.Visible = true;
            bt_EditarPRJ_Editar.Text = "Salvar";

            //LIMPA CAMPOS
            cb_LTFabrica_CadastroPRJ.Items.Clear();
            cb_LTFabrica_CadastroPRJ.Text = null;
            cb_LTFuncional_CadastroPRJ.Items.Clear();
            cb_LTFuncional_CadastroPRJ.Text = null;
            tb_Descricao_CadastroPRJ.Text = "";

            lb_Titulo_CadastarProjeto.Visible = false;
            lb_Titulo_EditarProjeto.Visible = true;
            pb_EditarPRJ.Visible = true;
            pb_CadastrarPRJ.Visible = false;

            limpaCadastroPRJ();

            #region CARREGA PROJETOS NO COMBOBOX

            if (bdConn.State == ConnectionState.Closed)
            {
                try
                {
                    //LIMPA COMBOBOX
                    cb_EditarPRJ_Projeto.Items.Clear();

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT cod_prj FROM projeto ORDER BY cod_prj;", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarPRJ_Projeto.Items.Add(dr["cod_prj"].ToString());

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }

            #endregion
        }

        //BOTÃO EDITAR USUÁRIOS - MENU PARAMETROS
        private void bt_EditarUsuario_Click(object sender, EventArgs e)
        {
            panel_EditarUser.Visible = true;
            panel_EditarUser.BringToFront();
            limpaEditarUsuario();

            #region ATUALIZA COMBOBOX DE ANALISTAS

            if (bdConn.State == ConnectionState.Closed)
            {
                try
                {
                    //LIMPA COMBOBOX
                    cb_EditarUser_Analista.Items.Clear();
                    cb_EditarUser_Analista.Items.Add("");

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT user_nome FROM usuarios", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarUser_Analista.Items.Add(dr["user_nome"].ToString());

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }
            #endregion
        }

        //BOTÃO CADASTRR RQF - MENU PARAMETROS
        private void bt_CadastrarRQF_Click(object sender, EventArgs e)
        {
            panel_CadastrarRQF_1.Visible = true;
            panel_CadastrarRQF_1.BringToFront();
        }

        //BOTÃO EDITAR RQF - MENU PARAMETROS
        private void bt_EditarRQF_Click(object sender, EventArgs e)
        {
            //HABILITA PANEL            
            panel_EditarRQF.Visible = true;
            panel_EditarRQF.BringToFront();

            limpaCampos_EditarRQF();
        }

        //BOTÃO NOVO STATUS - MENU PARAMETROS
        private void bt_NovoStatus_Click(object sender, EventArgs e)
        {
            //HABILITA PANEL            
            panel_NovoStatus1.Visible = true;
            panel_NovoStatus1.BringToFront();
            lb_Status.Text = "Adicionar Status";
            lb_Status2.Text = "Adicionar Status";
            bt_NS_Salvar.Text = "Salvar";
        }

        //BOTÃO NOVO STATUS - MENU PARAMETROS
        private void bt_EditarStatus_Click(object sender, EventArgs e)
        {
            //HABILITA PANEL            
            panel_NovoStatus1.Visible = true;
            panel_NovoStatus1.BringToFront();
            lb_Status.Text = "Editar Status";
            lb_Status2.Text = "Editar Status";
            bt_NS_Salvar.Text = "Salvar Alteração";
        }

        #region //****************************************** CADASTRO PROJETO ******************************************\\

        //BOTÃO SALVAR PROJETO
        private void bt_Salvar_CadastroPRJ_Click(object sender, EventArgs e)
        {
            string verificaCampos = vericaCampos_cadastroPRJ();
            if (verificaCampos == "")
            {
                try
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    //EXECUTA COMANDO
                    MySqlCommand command = new MySqlCommand("INSERT INTO projeto (cod_prj, lider_fabrica, lider_funcional, desc_prj, acompanhamento, acompanhamento_ext, acompanhamento_testes, status_prj) VALUES ('" + tb_Projeto_CadastroPRJ.Text + "','" + cb_LTFabrica_CadastroPRJ.Text + "','" + cb_LTFuncional_CadastroPRJ.Text + "','" + tb_Descricao_CadastroPRJ.Text + "', '" + ((tb_AddAcomp.Text == "") ? null : tb_AddAcomp.Text) + "', '" + ((tb_AddAcompExt.Text == "") ? null : tb_AddAcompExt.Text) + "', '" + ((tb_AddAcompTestes.Text == "") ? null : tb_AddAcompTestes.Text) + "', '" + cb_StatusPRJ.Text + "');", bdConn);
                    command.ExecuteNonQuery();

                    #region CADASTRA CRITICIDADE

                    command = new MySqlCommand("INSERT INTO prj_sla (cod_prj, criticidade_baixa, criticidade_media, criticidade_alta) VALUES ('" + tb_Projeto_CadastroPRJ.Text + "', '" + ud_SLA_Baixa.Text.Substring(0, 2) + "', '" + ud_SLA_Media.Text.Substring(0, 2) + "', '" + ud_SLA_Alta.Text.Substring(0, 2) + "');", bdConn);
                    command.ExecuteNonQuery();

                    #endregion

                    command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Cadastro de Projeto', '" + tb_Projeto_CadastroPRJ.Text + "');", bdConn);
                    command.ExecuteNonQuery();

                    //FECHA CONEXÃO
                    bdConn.Close();

                    MessageBox.Show("Projeto foi cadastrado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    limpaCadastroPRJ();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                MessageBox.Show(verificaCampos, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //VERIFICA CAMPOS DO PROJETO
        string vericaCampos_cadastroPRJ()
        {
            if (tb_Projeto_CadastroPRJ.Text == "PRJ00000000" && cb_EditarPRJ_Projeto.Text == null)
                return "Informe o Projeto!";

            if (cb_LTFabrica_CadastroPRJ.Text == "")
                return "Informe o Líder Fábrica!";

            if (cb_LTFuncional_CadastroPRJ.Text == "")
                return "Informe o Líder Funcional!";

            if (cb_StatusPRJ.Text == "")
                return "Informe o Status do Projeto!";

            if (tb_AddAcompExt.Text != "")
            {
                string[] emails = tb_AddAcompExt.Text.Split(';');
                foreach (string email in emails)
                {
                    string veemail = email.Replace(" ", "");
                    if (VeEmail(veemail) == false)
                        return "O email externo '" + email + "' não é válido!";
                }
            }

            return "";
        }

        //LIMPA CADASTRO DE PROJETO
        void limpaCadastroPRJ()
        {
            tb_Projeto_CadastroPRJ.Text = "00000000";
            cb_LTFabrica_CadastroPRJ.Text = null;
            cb_LTFuncional_CadastroPRJ.Text = null;
            tb_Descricao_CadastroPRJ.Text = "";
            tb_AddAcomp.Text = "";
            tb_AddAcompExt.Text = "";
            tb_AddAcompTestes.Text = "";
            ud_SLA_Baixa.Text = "01 Dia";
            ud_SLA_Media.Text = "01 Dia";
            ud_SLA_Alta.Text = "01 Dia";
            cb_StatusPRJ.Text = null;
        }

        //BOTÃO RETORNAR PARA PARAMENTROS
        private void bt_Return_CadastroPRJ_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //CONTROLE TEXTBOX DESCRIÇÃO PROJETO
        private void tb_Descricao_CadastroPRJ_TextChanged(object sender, EventArgs e)
        {
            lb_CadastroPRJ_controleDesc.Text = (500 - tb_Descricao_CadastroPRJ.TextLength).ToString();
        }

        //AJUDA ACOMPANHAMENTO
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Campo para adicionar usuários que estaram acompanhando as notificações do projeto.", "Acompanhar Projeto!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //ADICIONAR USUARIO NO ACOMPANHAMENTO
        private void bt_Add_AcompPRJ_Click(object sender, EventArgs e)
        {
            form_AddAcompanharPRJ formADD = new form_AddAcompanharPRJ();
            formADD.ShowDialog();
            if (formADD.nomeSelecionados != "")
                tb_AddAcomp.Text += ((tb_AddAcomp.Text == "") ? formADD.nomeSelecionados : (", " + formADD.nomeSelecionados));
        }

        //LIMPAR ACOMPANHAMENTO
        private void bt_LimpaAcomp_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Deseja tirar todas as pessoas do acompanhamento?", "Limpar", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
                tb_AddAcomp.Text = "";
        }

        //ADICIONA ; NO EMAIL EXTERNO
        private void tb_AddAcompExt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Space)
            {
                tb_AddAcompExt.Text = tb_AddAcompExt.Text + "; ";

                tb_AddAcompExt.SelectionStart = tb_AddAcompExt.Text.Length - 1; // add some logic if length is 0
                tb_AddAcompExt.SelectionLength = 0;
            }
        }

        //VERIFICA EMAIL
        public static bool VeEmail(string strEmail)
        {
            string strModelo = "^([0-9a-zA-Z]([-.\\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\\w]*[0-9a-zA-Z]\\.)+[a-zA-Z]{2,9})$";
            if (System.Text.RegularExpressions.Regex.IsMatch(strEmail, strModelo))
                return true;
            else
                return false;
        }

        #endregion

        #region //****************************************** EDITAR PROJETO ******************************************\\

        //BOTÃO EDITAR - EDITAR PROJETO
        private void bt_EditarPRJ_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                if (cb_EditarPRJ_Projeto.Text != "")
                {
                    string verificaCampos = vericaCampos_cadastroPRJ();
                    if (verificaCampos == "")
                    {
                        //ABRE CONEXÃO
                        bdConn.Open();

                        //EXECUTA COMANDO
                        MySqlCommand command = new MySqlCommand("UPDATE projeto SET lider_fabrica = '" + cb_LTFabrica_CadastroPRJ.Text + "', lider_funcional = '" + cb_LTFuncional_CadastroPRJ.Text + "', desc_prj = '" + tb_Descricao_CadastroPRJ.Text + "', acompanhamento = '" + ((tb_AddAcomp.Text == "") ? null : tb_AddAcomp.Text) + "', acompanhamento_ext = '" + ((tb_AddAcompExt.Text == "") ? null : tb_AddAcompExt.Text) + "', acompanhamento_testes = '" + ((tb_AddAcompTestes.Text == "") ? null : tb_AddAcompTestes.Text) + "', status_prj = '" + cb_StatusPRJ.Text + "' WHERE cod_prj = '" + cb_EditarPRJ_Projeto.Text + "';", bdConn);
                        command.ExecuteNonQuery();

                        command = new MySqlCommand("UPDATE prj_sla SET criticidade_baixa = " + ud_SLA_Baixa.Text.Substring(0, 2) + ", criticidade_media = " + ud_SLA_Media.Text.Substring(0, 2) + ", criticidade_alta = " + ud_SLA_Alta.Text.Substring(0, 2) + " WHERE cod_prj = '" + cb_EditarPRJ_Projeto.Text + "';", bdConn);
                        command.ExecuteNonQuery();

                        command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Editar Projeto', '" + tb_Projeto_CadastroPRJ.Text + "');", bdConn);
                        command.ExecuteNonQuery();

                        MessageBox.Show("Projeto foi alterado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        bt_Return_CadastroPRJ.PerformClick();
                    }
                    else
                        MessageBox.Show(verificaCampos, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                else
                    MessageBox.Show("Selecione o projeto que deseja alterar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //FECHA CONEXÃO
                bdConn.Close();
            }
        }

        //COMBOBOX PROJETOS - EDITAR PROJETO
        private void cb_EditarPRJ_Projeto_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //LIMPA COMBOBOX                
                cb_LTFabrica_CadastroPRJ.Items.Clear();
                cb_LTFabrica_CadastroPRJ.Text = null;
                cb_LTFuncional_CadastroPRJ.Items.Clear();
                cb_LTFuncional_CadastroPRJ.Text = null;
                tb_Descricao_CadastroPRJ.Text = "";

                //ABRE CONEXÃO
                bdConn.Open();

                #region ATUALIZA COMBOBOX DE ANALISTAS

                MySqlCommand command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Desenvolvimento';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                while (dr.Read())
                    cb_LTFabrica_CadastroPRJ.Items.Add(dr["user_nome"].ToString());
                dr.Close();

                command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Funcional';", bdConn);
                dr = command.ExecuteReader();
                while (dr.Read())
                    cb_LTFuncional_CadastroPRJ.Items.Add(dr["user_nome"].ToString());
                dr.Close();

                #endregion

                command = new MySqlCommand("SELECT * FROM projeto NATURAL JOIN prj_sla WHERE cod_prj = '" + cb_EditarPRJ_Projeto.Text + "';", bdConn);
                dr = command.ExecuteReader();
                if (dr.Read())
                {
                    cb_LTFabrica_CadastroPRJ.Text = dr["lider_fabrica"].ToString();
                    cb_LTFuncional_CadastroPRJ.Text = dr["lider_funcional"].ToString();
                    tb_Descricao_CadastroPRJ.Text = dr["desc_prj"].ToString();
                    tb_AddAcomp.Text = dr["acompanhamento"].ToString();
                    tb_AddAcompExt.Text = dr["acompanhamento_ext"].ToString();
                    tb_AddAcompTestes.Text = dr["acompanhamento_testes"].ToString();
                    ud_SLA_Baixa.Text = ((Int16.Parse(dr["criticidade_baixa"].ToString()) < 10) ? "0" : "") + dr["criticidade_baixa"].ToString() + ((dr["criticidade_baixa"].ToString() == "1") ? " Dia" : " Dias");
                    ud_SLA_Media.Text = ((Int16.Parse(dr["criticidade_media"].ToString()) < 10) ? "0" : "") + dr["criticidade_media"].ToString() + ((dr["criticidade_media"].ToString() == "1") ? " Dia" : " Dias");
                    ud_SLA_Alta.Text = ((Int16.Parse(dr["criticidade_alta"].ToString()) < 10) ? "0" : "") + dr["criticidade_alta"].ToString() + ((dr["criticidade_alta"].ToString() == "1") ? " Dia" : " Dias");
                    cb_StatusPRJ.Text = dr["status_prj"].ToString();
                }
                dr.Close();

                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO RETORNAR PARA PARAMENTROS - EDITAR PROJETO
        private void bt_Return_EditarPRJ_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        #endregion

        #region //****************************************** EDITAR USUÁRIO ******************************************\\

        //BOTÃO - ALTERAR USUÁRIO
        private void bt_EditarUser_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                //EXECUTA COMANDO
                MySqlCommand command = new MySqlCommand("UPDATE usuarios SET user_equipe = '" + cb_EditarUser_Equipe.Text + "', user_lvl = '" + ((cb_EditarUser_TipoUser.Text == "Comum") ? "0" : "1") + "', permissao_relatorio = '" + ((cb_PermissaoRelatorioOcorrencias.Checked == true) ? "sim" : "nao") + "' WHERE user_nome = '" + cb_EditarUser_Analista.Text + "';", bdConn);
                command.ExecuteNonQuery();

                command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Alterar Usuario', '" + cb_EditarUser_Analista.Text + "');", bdConn);
                command.ExecuteNonQuery();

                //FECHA CONEXÃO
                bdConn.Close();

                MessageBox.Show("Analista foi alterado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                limpaEditarUsuario();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //CARREGA INFORMAÇÕES DO ANALISTA
        private void cb_EditarUser_Analista_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_EditarUser_Analista.Text != "")
            {

                try
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT user_email, user_lvl, user_equipe, permissao_relatorio FROM usuarios WHERE user_nome = '" + cb_EditarUser_Analista.Text + "'", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    if (dr.Read())
                    {
                        tb_EditarUser_Email.Text = dr["user_email"].ToString();
                        cb_EditarUser_Equipe.Text = dr["user_equipe"].ToString();
                        cb_EditarUser_TipoUser.Text = ((dr["user_lvl"].ToString() == "0") ? "Comum" : "Administrador");
                        cb_PermissaoRelatorioOcorrencias.Checked = ((dr["permissao_relatorio"].ToString() == "sim") ? true : false);
                    }

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }


            }
            else
                limpaEditarUsuario();

        }

        //BOTÃO RETORNAR PARA PARAMENTROS
        private void bt_EditarUser_Return_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //LIMPAR USUÁRIO
        void limpaEditarUsuario()
        {
            cb_EditarUser_Analista.Text = null;
            tb_EditarUser_Email.Text = "";
            cb_EditarUser_Equipe.Text = null;
            cb_EditarUser_TipoUser.Text = null;
            cb_PermissaoRelatorioOcorrencias.Checked = false;
        }

        private void cb_EditarUser_TipoUser_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_EditarUser_TipoUser.Text == "Administrador")
                cb_PermissaoRelatorioOcorrencias.Checked = true;
            else
                cb_PermissaoRelatorioOcorrencias.Checked = false;
        }

        #endregion

        #region //****************************************** CADASTRAR RQF ******************************************\\

        //BOTÃO RETORNAR PARA PARAMENTROS- CADASTRO RQF
        private void bt_CadastrarRQF_Return_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //TEXTBOX PESQUISA PROJETOS - CADASTRO RQF
        private void tb_CadastrarRQF_SelecionarProjeto_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_CadastrarRQF_SelecionarProjeto.Text != "")
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    bdDataSet = new DataSet();
                    bdAdapter = new MySqlDataAdapter("SELECT cod_prj FROM projeto WHERE status_prj = 'Construção' AND cod_prj like '%" + tb_CadastrarRQF_SelecionarProjeto.Text + "%';", bdConn);
                    bdAdapter.Fill(bdDataSet, "projeto");
                    dataGrid_CadastrarRQF.DataSource = bdDataSet;

                    if (bdDataSet.Tables["projeto"].Rows.Count == 0)
                        lb_ProjetoNaoEncontrado_CadastrarRQF.Visible = true;
                    else
                        lb_ProjetoNaoEncontrado_CadastrarRQF.Visible = false;

                    dataGrid_CadastrarRQF.DataMember = "projeto";

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                else
                {
                    lb_ProjetoNaoEncontrado_CadastrarRQF.Visible = false;
                    if (this.dataGrid_CadastrarRQF.DataSource != null)
                        this.dataGrid_CadastrarRQF.DataSource = null;
                    else
                    {
                        this.dataGrid_CadastrarRQF.Rows.Clear();
                        this.dataGrid_CadastrarRQF.Columns.Clear();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO OK - SELECIONAR PROJETO - CADASTRO RQF
        private void bt_CadastrarRQF_OK_Click(object sender, EventArgs e)
        {
            if (dataGrid_CadastrarRQF.CurrentRow != null)
                if (bdDataSet.Tables["projeto"].Rows.Count > 0)
                {
                    try
                    {
                        cb_CadastroRQF_RespFuncional.Items.Clear();
                        cb_CadastroRQF_RespFuncional.Items.Add("");

                        bdConn.Open();
                        MySqlCommand command = new MySqlCommand("select user_nome from usuarios where user_equipe = 'Fábrica Funcional';", bdConn);
                        MySqlDataReader dr = command.ExecuteReader();
                        while (dr.Read())
                            cb_CadastroRQF_RespFuncional.Items.Add(dr["user_nome"].ToString());
                        dr.Close();
                        bdConn.Close();

                        tb_CadastroRQF_Projeto.Text = dataGrid_CadastrarRQF.CurrentRow.Cells[0].Value.ToString();

                        panel_CadastrarRQF_1.Visible = false;
                        panel_CadastrarRQF_2.Visible = true;
                        panel_CadastrarRQF_2.BringToFront();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        bdConn.Close();
                    }
                }
                else
                    MessageBox.Show("Nenhum projeto foi encontrado! Pesquise novamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //CONTROLE MASCARA DO TEXTBOX - CADASTRO RQF
        private void rb_RQF_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_RQF.Checked == true)
            {
                tb_CadastroRQF_RQF.Mask = "RQF999";
            }
            else
                tb_CadastroRQF_RQF.Mask = "RQNF999";
        }

        //CONTROLE TEXTBOX DESCRIÇÃO - CADASTRO RQF
        private void tb_CadastroRQF_Descricao_TextChanged(object sender, EventArgs e)
        {
            lb_CadastroRQF_ControleDescricao.Text = (300 - tb_CadastroRQF_Descricao.TextLength).ToString();
        }

        //BOTÃO CADASTRAR - CADASTRO RQF
        private void bt_CadastroRQF_Cadastrar_Click(object sender, EventArgs e)
        {
            if (verificaCampos_cadastroRQF())
            {
                try
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    //EXECUTA COMANDO
                    MySqlCommand command = new MySqlCommand("INSERT INTO df_rqf (cod_rqf, cod_prj, resp_funcional, desc_rqf) VALUES ('" + tb_CadastroRQF_RQF.Text + "', '" + tb_CadastroRQF_Projeto.Text + "', '" + cb_CadastroRQF_RespFuncional.Text + "', '" + tb_CadastroRQF_Descricao.Text + "');", bdConn);
                    command.ExecuteNonQuery();

                    //ATUALIZA OCORRENCIAS COM A RQF
                    command = new MySqlCommand("UPDATE ocorrencia SET atrib_para = '" + cb_CadastroRQF_RespFuncional.Text + "' WHERE cod_prj = '" + tb_CadastroRQF_Projeto.Text + "' AND cod_rqf = '" + tb_CadastroRQF_RQF.Text + "';", bdConn);
                    command.ExecuteNonQuery();

                    command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Cadastro de RQF', '" + tb_CadastroRQF_RQF.Text + " foi cadastrada para o projeto " + tb_CadastroRQF_Projeto.Text + " com o analista " + cb_CadastroRQF_RespFuncional.Text + " como RF ');", bdConn);
                    command.ExecuteNonQuery();

                    //FECHA CONEXÃO
                    bdConn.Close();

                    //EXIBE MENSSAGEM
                    MessageBox.Show("RQF cadastrada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    //LIMPA CAMPOS
                    limpaCadastroRQF();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Já existe a " + tb_CadastroRQF_RQF.Text + " cadastrada para o projeto " + tb_CadastroRQF_Projeto.Text + ".", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    bdConn.Close();
                }
            }
            else
                MessageBox.Show("Preencha os campos corretamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);


        }

        //VERIFICA CAMPOS - CADASTRO RQF
        bool verificaCampos_cadastroRQF()
        {
            if (tb_CadastroRQF_RQF.Text == "RQF" || tb_CadastroRQF_RQF.Text == "RQNF")
                return false;

            return true;
        }

        //LIMPA CAMPOS - CADASTRO RQF
        void limpaCadastroRQF()
        {
            tb_CadastroRQF_RQF.Text = "";
            cb_CadastroRQF_RespFuncional.Text = null;
            tb_CadastroRQF_Descricao.Text = "";
        }

        //DUPLO CLICK NA CELULA CELECIONADA - CADASTRO RQF
        private void dataGrid_CadastrarRQF_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGrid_CadastrarRQF.CurrentRow != null)
                bt_CadastrarRQF_OK.PerformClick();
            else
                MessageBox.Show("Selecionar um projeto da lista antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //BOTÃO RETURN - CADASTRO RQF 2
        private void bt_CadastroRQF_Return_Click(object sender, EventArgs e)
        {
            panel_CadastrarRQF_2.Visible = false;
            panel_CadastrarRQF_1.Visible = true;
            panel_CadastrarRQF_1.BringToFront();

            tb_CadastroRQF_Projeto.Text = "";
            tb_CadastrarRQF_SelecionarProjeto.Text = "";
        }

        #endregion

        #region //****************************************** EDITAR RQF ******************************************\\

        //BOTÃO EDITAR - EDITAR RQF
        private void bt_EditarRQF_Editar_Click(object sender, EventArgs e)
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                //EXECUTA COMANDO
                MySqlCommand command = new MySqlCommand("UPDATE df_rqf SET resp_funcional = '" + cb_EditarRQF_RespFuncional.Text + "', desc_rqf = '" + tb_EditarRQF_DescricaoRQF.Text + "' WHERE cod_prj = '" + cb_EditarRQF_Projeto.Text + "' AND cod_rqf = '" + cb_EditarRQF_RQF.Text + "';", bdConn);
                command.ExecuteNonQuery();

                command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Editar RQF', '" + cb_EditarRQF_RQF.Text + " foi editada. Responsavel Funcinal: " + cb_EditarRQF_RespFuncional.Text + ". Descricao: " + tb_EditarRQF_DescricaoRQF.Text + "');", bdConn);
                command.ExecuteNonQuery();

                //FECHA CONEXÃO
                bdConn.Close();

                MessageBox.Show("RQF foi alterada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                limpaCampos_EditarRQF();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //LIMPA CAMPOS - EDITAR RQF
        void limpaCampos_EditarRQF()
        {
            #region RECUPERA PROJETOS COM RQF CADASTRADAS
            if (bdConn.State == ConnectionState.Closed)
            {
                try
                {
                    //LIMPA COMBOBOX
                    cb_EditarRQF_Projeto.Items.Clear();

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT distinct cod_prj FROM df_rqf ORDER BY cod_prj;", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarRQF_Projeto.Items.Add(dr["cod_prj"].ToString());

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }
            #endregion

            //LIMPA COMBOBOX
            cb_EditarRQF_RQF.Items.Clear();
            cb_EditarRQF_RespFuncional.Items.Clear();
            tb_EditarRQF_DescricaoRQF.Text = "";
        }

        //BOTÃO RETORNAR PARA PARAMENTROS - EDITAR RQF
        private void bt_EditarRQF_Return_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //TEXTBOX PROJETOS  - EDITAR RQF
        private void cb_EditarRQF_Projeto_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //LIMPA COMBOBOX                
                cb_EditarRQF_RQF.Items.Clear();
                cb_EditarRQF_RQF.Text = null;
                cb_EditarRQF_RespFuncional.Items.Clear();
                cb_EditarRQF_RespFuncional.Text = null;
                tb_EditarRQF_DescricaoRQF.Text = "";

                //ABRE CONEXÃO
                bdConn.Open();

                MySqlCommand command = new MySqlCommand("SELECT cod_rqf FROM df_rqf WHERE cod_prj = '" + cb_EditarRQF_Projeto.Text + "' ORDER BY cod_rqf;", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                while (dr.Read())
                    cb_EditarRQF_RQF.Items.Add(dr["cod_rqf"].ToString());
                dr.Close();

                command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Funcional' ORDER BY user_nome;", bdConn);
                dr = command.ExecuteReader();
                while (dr.Read())
                    cb_EditarRQF_RespFuncional.Items.Add(dr["user_nome"].ToString());
                dr.Close();

                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //TEXTBOX RQFs - EDITAR RQF
        private void cb_EditarRQF_RQF_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                MySqlCommand command = new MySqlCommand("SELECT resp_funcional, desc_rqf FROM df_rqf WHERE cod_prj = '" + cb_EditarRQF_Projeto.Text + "' AND cod_rqf = '" + cb_EditarRQF_RQF.Text + "';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                if (dr.Read())
                {
                    cb_EditarRQF_RespFuncional.Text = dr["resp_funcional"].ToString();
                    tb_EditarRQF_DescricaoRQF.Text = dr["desc_rqf"].ToString();
                }
                dr.Close();

                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        #endregion

        #region //****************************************** NOVO/EDITAR STATUS ******************************************\\

        #region NOVO STATUS 1

        //BOTÃO RETORNAR PARA PARAMENTROS - NOVO STATUS
        private void bt_NS1_Return_Click(object sender, EventArgs e)
        {
            inicioMenuParametros();
        }

        //BOTÃO SELECIONA STATUS - NOVO STATUS
        private void bt_NS_Continuar_Click(object sender, EventArgs e)
        {
            //HABILITA ADICIONAR EQUIPE PARA STATUS
            gb_PermissaoNS.Enabled = true;

            //VERIFICA TIPO DE STATUS
            if (rb_NS_StatusGeral.Checked == true)
            {
                tb_NS_TipoStatus.Text = rb_NS_StatusGeral.Text;
                gb_NS_AualizaStatusGeral.Enabled = false;
                gb_NS_EnviaEmail.Enabled = false;
                gb_NS_EncerrarO.Enabled = true;
                rb_NS_Nao_Encerra.Checked = true;
                rb_Testes_Nao.Checked = true;
                gb_Testes.Enabled = false;

                if (lb_Status.Text == "Adicionar Status")
                {
                    tb_NS_NomeStatus.Visible = true;
                    cb_EditarStatus.Visible = false;
                    bt_ES_Excluir.Visible = false;
                }
                else
                {
                    tb_NS_NomeStatus.Visible = false;
                    cb_EditarStatus.Visible = true;
                    bt_ES_Excluir.Visible = true;

                    cb_EditarStatus.Items.Clear();

                    bdConn.Open();
                    MySqlCommand command = new MySqlCommand("SELECT nome_status FROM status_geral ORDER BY nome_status;", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarStatus.Items.Add(dr[0].ToString());
                    dr.Close();
                    bdConn.Close();
                }
            }
            else if (rb_NS_StatusResposta.Checked == true)
            {
                tb_NS_TipoStatus.Text = rb_NS_StatusResposta.Text;
                gb_NS_EnviaEmail.Enabled = true;
                gb_NS_AualizaStatusGeral.Enabled = true;
                gb_NS_EncerrarO.Enabled = false;
                gb_Testes.Enabled = false;
                rb_Testes_Nao.Checked = true;
                carregaStatusGeral_NS();

                if (lb_Status.Text == "Adicionar Status")
                {
                    tb_NS_NomeStatus.Visible = true;
                    cb_EditarStatus.Visible = false;
                    bt_ES_Excluir.Visible = false;
                }
                else
                {
                    tb_NS_NomeStatus.Visible = false;
                    cb_EditarStatus.Visible = true;
                    bt_ES_Excluir.Visible = true;

                    cb_EditarStatus.Items.Clear();

                    bdConn.Open();
                    MySqlCommand command = new MySqlCommand("SELECT nome_status FROM status_resposta ORDER BY nome_status;", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarStatus.Items.Add(dr[0].ToString());
                    dr.Close();
                    bdConn.Close();
                }

            }
            else if (rb_NS_Impacto.Checked == true)
            {

                tb_NS_TipoStatus.Text = rb_NS_Impacto.Text;
                gb_PermissaoNS.Enabled = false;
                gb_NS_EnviaEmail.Enabled = false;
                gb_NS_AualizaStatusGeral.Enabled = false;
                gb_NS_EncerrarO.Enabled = false;
                gb_Testes.Enabled = true;
                rb_Testes_Nao.Checked = true;

                if (lb_Status.Text == "Adicionar Status")
                {
                    tb_NS_NomeStatus.Visible = true;
                    cb_EditarStatus.Visible = false;
                    bt_ES_Excluir.Visible = false;
                }
                else
                {
                    tb_NS_NomeStatus.Visible = false;
                    cb_EditarStatus.Visible = true;
                    bt_ES_Excluir.Visible = true;

                    cb_EditarStatus.Items.Clear();

                    bdConn.Open();
                    MySqlCommand command = new MySqlCommand("SELECT nome_status FROM status_impacto ORDER BY nome_status;", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    while (dr.Read())
                        cb_EditarStatus.Items.Add(dr[0].ToString());
                    dr.Close();
                    bdConn.Close();
                }
            }

            //HABILITA PANEL            
            panel_NovoStatus2.Visible = true;
            panel_NovoStatus2.BringToFront();
        }

        void carregaStatusGeral_NS()
        {
            try
            {
                cb_NS_AtualizaStatus_StatusGeral.Items.Clear();
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT nome_status FROM status_geral", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                while (dr.Read())
                    cb_NS_AtualizaStatus_StatusGeral.Items.Add(dr["nome_status"].ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Detalhes: " + ex.Message, "Erro load Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                bdConn.Close();
            }
        }

        #endregion

        #region NOVO STATUS 2

        //BOTÃO RETORNAR NS1 - NOVO STATUS
        private void bt_NS2_Return_Click(object sender, EventArgs e)
        {
            limpaNS();

            panel_NovoStatus2.Visible = false;
        }

        //BOTÃO SALVAR - NOVO STATUS
        private void bt_NS_Salvar_Click(object sender, EventArgs e)
        {
            if (lb_Status2.Text == "Adicionar Status")
            {
                if (tb_NS_TipoStatus.Text != "Impacto")
                    if ((tb_NS_NomeStatus.Text != "") && (cb_NS_Fabrica.Checked != false || cb_NS_Funcional.Checked != false))
                        salvaNovoStatus();
                    else
                        MessageBox.Show("Preencha os campos corretamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    if (tb_NS_NomeStatus.Text != "")
                    salvaNovoStatus();
                else
                    MessageBox.Show("Preencha os campos corretamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                editarStatus();
            }

        }

        //LIMPA NOVO STATUS
        void limpaNS()
        {
            tb_NS_NomeStatus.Text = "";
            cb_NS_Fabrica.Checked = false;
            cb_NS_Funcional.Checked = false;
        }

        //SALVA NOVO STATUS
        void salvaNovoStatus()
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                string queryNS = "";
                string VALUES = "";


                if (tb_NS_TipoStatus.Text == "Status Geral")
                {
                    VALUES = "'" + tb_NS_NomeStatus.Text + "', '" + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "', '" + ((rb_NS_Sim_Encerra.Checked == true) ? "sim" : "nao") + "'";
                    queryNS = "INSERT INTO status_geral VALUES (" + VALUES + ");";
                }
                else if (tb_NS_TipoStatus.Text == "Status Resposta")
                {
                    if (rb_NS_AtualizaStatus_Sim.Checked)
                        if (cb_NS_AtualizaStatus_StatusGeral.Text != "")
                        {
                            VALUES = "'" + tb_NS_NomeStatus.Text + "', '" + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "', '" + ((rb_NS_EnviarEmail_Sim.Checked == true) ? "sim" : "nao") + "', '" + (cb_NS_AtualizaStatus_StatusGeral.Text) + "'";
                            queryNS = "INSERT INTO status_resposta VALUES (" + VALUES + ");";
                        }
                        else
                            MessageBox.Show("Selecionar o nome do Status Geral para ser configurado como atualização automatica!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        VALUES = "'" + tb_NS_NomeStatus.Text + "', '" + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "', '" + ((rb_NS_EnviarEmail_Sim.Checked == true) ? "sim" : "nao") + "'";
                        queryNS = "INSERT INTO status_resposta (nome_status, equipe, envia_email) VALUES (" + VALUES + ");";
                    }

                }
                else if (tb_NS_TipoStatus.Text == "Impacto")
                {
                    queryNS = "INSERT INTO status_impacto (nome_status, email_testes) VALUES ('" + tb_NS_NomeStatus.Text + "', '" + ((rb_Testes_Sim.Checked == true) ? "Sim" : "Nao") + "');";
                }


                //EXECUTA COMANDO
                MySqlCommand command = new MySqlCommand(queryNS, bdConn);
                command.ExecuteNonQuery();

                command = new MySqlCommand("INSERT INTO log_alt (analista, tipo_alt, acao_alt) VALUES ('" + lb_UserON.Text + "', 'Criar Status', 'Foi criado o Status " + tb_NS_NomeStatus.Text + ", com permissao para " + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "');", bdConn);
                command.ExecuteNonQuery();

                //FECHA CONEXÃO
                bdConn.Close();

                //EXIBE MENSSAGEM
                MessageBox.Show("Novo Status criado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //LIMPA CAMPOS
                limpaNS();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                bdConn.Close();
            }
        }

        //EDITAR STATUS
        void editarStatus()
        {
            switch (tb_NS_TipoStatus.Text)
            {
                case "Status Geral":
                    if (cb_EditarStatus.Text != "")
                    {
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("UPDATE status_geral SET equipe = '" + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "', encerra_ocorrencia = '" + ((rb_NS_Sim_Encerra.Checked == true) ? "sim" : "nao") + "' WHERE nome_status = '" + cb_EditarStatus.Text + "';", bdConn);
                            command.ExecuteNonQuery();

                            MessageBox.Show("O Status Geral (" + cb_EditarStatus.Text + ") foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            bt_NS2_Return.PerformClick();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                    }
                    else
                        MessageBox.Show("Selecionar um Status para Editar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case "Status Resposta":
                    if (cb_EditarStatus.Text != "")
                    {
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("UPDATE status_resposta SET equipe = '" + (((cb_NS_Fabrica.Checked == true) ? "FD" : "") + ((cb_NS_Funcional.Checked == true) ? "FF" : "")) + "', envia_email = '" + ((rb_NS_EnviarEmail_Sim.Checked == true) ? "sim" : "nao") + "', status_automatico_sg = '" + (cb_NS_AtualizaStatus_StatusGeral.Text) + "' WHERE nome_status = '" + cb_EditarStatus.Text + "';", bdConn);
                            command.ExecuteNonQuery();

                            MessageBox.Show("O Status Resposta (" + cb_EditarStatus.Text + ") foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            bt_NS2_Return.PerformClick();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                    }
                    else
                        MessageBox.Show("Selecionar um Status para Editar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
                case "Impacto":
                    if (cb_EditarStatus.Text != "")
                    {
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("UPDATE status_impacto SET email_testes = '" + ((rb_Testes_Sim.Checked == true) ? "Sim" : "Nao") + "' WHERE nome_status = '" + cb_EditarStatus.Text + "';", bdConn);
                            command.ExecuteNonQuery();

                            MessageBox.Show("O Status Impacto (" + cb_EditarStatus.Text + ") foi atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            bt_NS2_Return.PerformClick();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status de Impacto", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                    }
                    else
                        MessageBox.Show("Selecionar um Status para Editar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    break;
            }
        }

        //BOTÃO EXCLUIR
        private void bt_ES_Excluir_Click(object sender, EventArgs e)
        {
            if (cb_EditarStatus.Text != "")
            {
                DialogResult result = MessageBox.Show("Deseja excluir o " + tb_NS_TipoStatus.Text + " (" + cb_EditarStatus.Text + ")?", "Excluir", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    try
                    {
                        string table = "";
                        switch (tb_NS_TipoStatus.Text)
                        {
                            case "Status Geral":
                                table = "status_geral";
                                break;
                            case "Status Resposta":
                                table = "status_resposta";
                                break;
                            case "Impacto":
                                table = "status_impacto";
                                break;
                        }

                        bdConn.Open();
                        MySqlCommand command = new MySqlCommand("DELETE FROM " + table + " WHERE nome_status = '" + cb_EditarStatus.Text + "';", bdConn);
                        command.ExecuteNonQuery();

                        MessageBox.Show(tb_NS_TipoStatus.Text + " (" + cb_EditarStatus.Text + ") excluído com sucesso.", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        bt_NS2_Return.PerformClick();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Erro ao excluir status!\n\nDescrição: " + ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        bdConn.Close();
                    }
                }
            }
            else
            {
                MessageBox.Show("Selecionar o Status que deseja excluir!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //RADIO BUTON ATUALIZA STATUS
        private void rb_NS_AtualizaStatus_Sim_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_NS_AtualizaStatus_Sim.Checked == true)
                cb_NS_AtualizaStatus_StatusGeral.Enabled = true;
            else
            {
                cb_NS_AtualizaStatus_StatusGeral.Enabled = false;
                cb_NS_AtualizaStatus_StatusGeral.Text = null;
            }

        }

        //SELECT COMBOBOX STATUS
        private void cb_EditarStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_EditarStatus.Text != "" || cb_EditarStatus.Text != null)
            {
                switch (tb_NS_TipoStatus.Text)
                {
                    case "Status Geral":
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("SELECT equipe, encerra_ocorrencia FROM status_geral WHERE nome_status = '" + cb_EditarStatus.Text + "'", bdConn);
                            MySqlDataReader dr = command.ExecuteReader();
                            if (dr.Read())
                            {
                                switch (dr["equipe"].ToString())
                                {
                                    case "FD":
                                        cb_NS_Fabrica.Checked = true;
                                        cb_NS_Funcional.Checked = false;
                                        break;
                                    case "FF":
                                        cb_NS_Fabrica.Checked = false;
                                        cb_NS_Funcional.Checked = true;
                                        break;
                                    default:
                                        cb_NS_Fabrica.Checked = true;
                                        cb_NS_Funcional.Checked = true;
                                        break;
                                }

                                if (dr["encerra_ocorrencia"].ToString() == "sim")
                                    rb_NS_Sim_Encerra.Checked = true;
                                else
                                    rb_NS_Nao_Encerra.Checked = true;
                            }
                            dr.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                        break;
                    case "Status Resposta":
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("SELECT equipe, envia_email, status_automatico_sg FROM status_resposta WHERE nome_status = '" + cb_EditarStatus.Text + "'", bdConn);
                            MySqlDataReader dr = command.ExecuteReader();
                            if (dr.Read())
                            {
                                switch (dr["equipe"].ToString())
                                {
                                    case "FD":
                                        cb_NS_Fabrica.Checked = true;
                                        cb_NS_Funcional.Checked = false;
                                        break;
                                    case "FF":
                                        cb_NS_Fabrica.Checked = false;
                                        cb_NS_Funcional.Checked = true;
                                        break;
                                    default:
                                        cb_NS_Fabrica.Checked = true;
                                        cb_NS_Funcional.Checked = true;
                                        break;
                                }

                                if (dr["envia_email"].ToString() == "sim")
                                    rb_NS_EnviarEmail_Sim.Checked = true;
                                else
                                    rb_NS_EnviarEmail_Nao.Checked = true;

                                if (dr["status_automatico_sg"].ToString() != "")
                                {
                                    rb_NS_AtualizaStatus_Sim.Checked = true;
                                    cb_NS_AtualizaStatus_StatusGeral.Text = dr["status_automatico_sg"].ToString();
                                }
                                else
                                {
                                    rb_NS_AtualizaStatus_Nao.Checked = true;
                                    cb_NS_AtualizaStatus_StatusGeral.Text = null;
                                }
                            }
                            dr.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                        break;
                    case "Impacto":
                        try
                        {
                            bdConn.Open();
                            MySqlCommand command = new MySqlCommand("SELECT email_testes FROM status_impacto  WHERE nome_status = '" + cb_EditarStatus.Text + "'", bdConn);
                            MySqlDataReader dr = command.ExecuteReader();
                            if (dr.Read())
                            {
                                if (dr["email_testes"].ToString() == "Sim")
                                    rb_Testes_Sim.Checked = true;
                                else
                                    rb_Testes_Nao.Checked = true;
                            }
                            dr.Close();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Detalhes: " + ex.Message, "Erro editar Status de Impacto", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            bdConn.Close();
                        }
                        break;
                }
            }
        }

        //BOTÃO AJUDA - ENCERRA OCORRENCIA
        private void pb_Ajuda_EncerraOcorrencia_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Se selecionada a opção sim, quando o Status Geral for atualizado para (" + tb_NS_NomeStatus.Text + ") a Ocorrência do mesmo será encerrada.\n\n Não sendo posível alterar status e adicionar respostas.", "Encerrar Ocorrência", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //BOTÃO AJUDA - ENVIA EMAIL
        private void pb_Ajuda_EnviaEmail_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Se selecionada a opção sim, quando o Status Resposta for atualizado para (" + cb_EditarStatus.Text + ") o Identificador da Ocorrência receberá um email notificando a atualização do status!", "Enviar Email", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //BOTÃO AJUDA - ATUALIZA STATUS
        private void pb_Ajuda_AtualizaStatus_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Se selecionada a opção sim, quando o Status Resposta for atualizado para (" + cb_EditarStatus.Text + "), automaticamente o Status Geral será atualizado.", "Atualizar Automaticamente o Status Geral", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion

        #endregion

        #endregion

        #region //****************************************** REGISTRO DE OCORRÊNCIA  ******************************************\\

        //INICIO REGISTRO DE OCORRENCIA
        void inicioRegistroOcorrencia()
        {
            panel_RegOcorrencia_1.BringToFront();
            panel_RegOcorrencia_2.SendToBack();
            panel_RegOcorrencia_3.SendToBack();

            panel_RegOcorrencia_1.Dock = DockStyle.Fill;
            panel_RegOcorrencia_2.Dock = DockStyle.Fill;
            panel_RegOcorrencia_3.Dock = DockStyle.Fill;

            panel_RegOcorrencia_1.Visible = true;
            panel_RegOcorrencia_2.Visible = false;
            panel_RegOcorrencia_3.Visible = false;

            tb_RegOcorrencia_SelecionaPRJ.Text = "";

            tb_RegOcorrencia_Projeto.Text = "";
            tb_RegOcorrencia_Identificador.Text = "";
            cb_RegOcorrencia_Sistema.Text = null;
            cb_RegOcorrencia_Classificacao.Text = null;
            tb_RegOcorrencia_RQF.Text = "";
            tb_RegOcorrencia_AtribuidoP.Text = "";

            tb_RegOcorrencia_Desc.Text = "";
            tb_RegOcorrencia_Quest.Text = "";
            tb_RegOcorrencia_Sugestao.Text = "";

            countAnexo = 0;
            caminhoOrigemA1 = "";
            caminhoOrigemA2 = "";
            caminhoOrigemA3 = "";
            caminhoOrigemA4 = "";
            caminhoOrigemA5 = "";
            lb_Anexos.Visible = false;
            link_Anexo1.Visible = false;
            link_Anexo2.Visible = false;
            link_Anexo3.Visible = false;
            link_Anexo4.Visible = false;
            link_Anexo5.Visible = false;

        }

        #region //****************************************** PANEL 1 ******************************************\\

        //TEXTBOX PESQUISA PROJETOS - REGISTRO OCORRENCIA
        private void tb_RegOcorrencia_SelecionaPRJ_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_RegOcorrencia_SelecionaPRJ.Text != "")
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    bdDataSet = new DataSet();
                    bdAdapter = new MySqlDataAdapter("SELECT cod_prj FROM projeto WHERE status_prj = 'Construção' AND cod_prj like '%" + tb_RegOcorrencia_SelecionaPRJ.Text + "%';", bdConn);
                    bdAdapter.Fill(bdDataSet, "projeto");
                    dataGrid_RegOcorrencia.DataSource = bdDataSet;

                    if (bdDataSet.Tables["projeto"].Rows.Count == 0)
                        lb_RegOcorrencia_NotFound.Visible = true;
                    else
                        lb_RegOcorrencia_NotFound.Visible = false;

                    dataGrid_RegOcorrencia.DataMember = "projeto";

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                else
                {
                    lb_RegOcorrencia_NotFound.Visible = false;
                    if (this.dataGrid_RegOcorrencia.DataSource != null)
                        this.dataGrid_RegOcorrencia.DataSource = null;
                    else
                    {
                        this.dataGrid_RegOcorrencia.Rows.Clear();
                        this.dataGrid_RegOcorrencia.Columns.Clear();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO OK - SELECIONAR PROJETO - REGISTOR DE OCORRENCIA
        private void bt_RegOcorrencia_OK_Click(object sender, EventArgs e)
        {
            if (dataGrid_RegOcorrencia.CurrentRow != null)
                if (bdDataSet.Tables["projeto"].Rows.Count > 0)
                {
                    try
                    {

                        tb_RegOcorrencia_Projeto.Text = dataGrid_RegOcorrencia.CurrentRow.Cells[0].Value.ToString();
                        tb_RegOcorrencia_Identificador.Text = lb_UserON.Text;

                        panel_RegOcorrencia_2.Visible = true;
                        panel_RegOcorrencia_2.BringToFront();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        bdConn.Close();
                    }
                }
                else
                    MessageBox.Show("Nenhum projeto foi encontrado! Pesquise novamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //DUPLO CLICK NA CELULA CELECIONADA - REGISTOR DE OCORRENCIA
        private void dataGrid_RegOcorrencia_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGrid_RegOcorrencia.CurrentRow != null)
                bt_RegOcorrencia_OK.PerformClick();
            else
                MessageBox.Show("Selecionar um projeto da lista antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        #endregion

        #region //****************************************** PANEL 2 ******************************************\\

        //BOTÃO RETURN - PANEL 2 - REGISTOR DE OCORRENCIA
        private void bt_RegOcorrencia2_Return_Click(object sender, EventArgs e)
        {
            inicioRegistroOcorrencia();
        }

        //BOTÃO NEXT - PANEL 2 - REGISTOR DE OCORRENCIA
        private void bt_RegOcorrencia2_Next_Click(object sender, EventArgs e)
        {
            if (verificaCampos_RegOcorrencia_2())
            {
                panel_RegOcorrencia_3.BringToFront();

                panel_RegOcorrencia_2.Visible = false;
                panel_RegOcorrencia_3.Visible = true;
            }

        }

        //CONTROLE TEXTBOX RQF - REGISTOR DE OCORRENCIA 
        private void tb_RegOcorrencia_RQF_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                MySqlCommand command = new MySqlCommand("SELECT resp_funcional FROM df_rqf WHERE cod_prj = '" + tb_RegOcorrencia_Projeto.Text + "' AND cod_rqf = '" + tb_RegOcorrencia_RQF.Text + "';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                if (dr.Read())
                {
                    if (dr["resp_funcional"].ToString() != "")
                    {
                        tb_RegOcorrencia_AtribuidoP.Text = dr["resp_funcional"].ToString();
                        tb_RegOcorrencia_AtribuidoP.Visible = true;
                        lb_RegOcorrencia_AtribuidoP.Visible = true;
                    }
                }
                else
                {
                    tb_RegOcorrencia_AtribuidoP.Text = "";
                    tb_RegOcorrencia_AtribuidoP.Visible = false;
                    lb_RegOcorrencia_AtribuidoP.Visible = false;
                }
                dr.Close();

                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //CONTROLE RADIOBUTTON RQF - REGISTOR DE OCORRENCIA 
        private void rb_RO_RQF_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_RO_RQF.Checked == true)
                tb_RegOcorrencia_RQF.Mask = "RQF999";
            else
                tb_RegOcorrencia_RQF.Mask = "RQNF999";
        }

        //VERIFICA CAMPOS REGISTRO DE OCORRENCIA 2
        bool verificaCampos_RegOcorrencia_2()
        {
            if (cb_RegOcorrencia_Sistema.Text == "")
            {
                MessageBox.Show("Selecionar um (Sistema) para antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (cb_RegOcorrencia_Classificacao.Text == "")
            {
                MessageBox.Show("Selecionar a (Classificação) da ocorrência antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (tb_RegOcorrencia_RQF.Text == "RQF" || tb_RegOcorrencia_RQF.Text == "RQNF")
            {
                MessageBox.Show("Digitar a RQF/RQNF antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            else if (verificaEspacoVazio(tb_RegOcorrencia_RQF.Text))
            {
                MessageBox.Show("RQF/RQNF é inválida.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (cb_RegOcorrencia_Criticidade.Text == "")
            {
                MessageBox.Show("Selecionar o (Impacto na Construção da RQF/RQNF) antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        //VERIFICA ESPAÇO VAZIO
        bool verificaEspacoVazio(string palavra)
        {
            string[] quebraPalavra = palavra.Split(' ');
            if (quebraPalavra.Length > 1)
                return true;
            return false;
        }

        #endregion

        #region //****************************************** PANEL 3 ******************************************\\

        //BOTÃO REGISTRAR - REGISTAR - REGISTOR DE OCORRENCIA
        private void bt_RegOcorrencia3_Registrar_Click(object sender, EventArgs e)
        {

            if (verificaCampos_RegOcorrencia())
            {
                try
                {
                    string atribuidoPara = "";
                    string impacto_cttu = "";
                    string COD_OCCORENCIA = "";
                    string pastaAnexo = @"\\192.168.10.3\Acc_Oi_BH\Troca\Controle de Ocorrências - Anexos";
                    string pastaAnexoOco = "";

                    //ABRE CONEXÃO
                    bdConn.Open();
                    MySqlCommand command;
                    MySqlDataReader dr;

                    //CODIGO OCORRENCIA
                    #region CRIA CODIGO DA OCORRENCIA

                    command = new MySqlCommand("SELECT count(cod_oco) FROM ocorrencia WHERE cod_prj = '" + tb_RegOcorrencia_Projeto.Text + "';", bdConn);
                    dr = command.ExecuteReader();
                    if (dr.Read())
                        COD_OCCORENCIA = "P" + (tb_RegOcorrencia_Projeto.Text.Substring(6, 5)) + "-OC" + (Int16.Parse(dr["count(cod_oco)"].ToString()) + 1).ToString();
                    dr.Close();

                    #endregion

                    //RESPONSÁVEL PELA RQF
                    #region VERIFICA RESPONSÁVEL PELA RQF
                    if (!String.IsNullOrEmpty(tb_RegOcorrencia_AtribuidoP.Text))
                        atribuidoPara = tb_RegOcorrencia_AtribuidoP.Text;
                    else
                    {
                        command = new MySqlCommand("SELECT lider_funcional FROM projeto WHERE cod_prj = '" + tb_RegOcorrencia_Projeto.Text + "';", bdConn);
                        dr = command.ExecuteReader();
                        if (dr.Read())
                            atribuidoPara = dr["lider_funcional"].ToString();
                        dr.Close();
                    }
                    #endregion

                    //VERIFICA CRITICIDADE
                    #region VERIFICA CRITICIDADE
                    if (Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) <= 40 && Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) != 10)
                        impacto_cttu = "Baixa";
                    else if (Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) <= 80 && Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) > 40)
                        impacto_cttu = "Média";
                    else if (Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) > 80 || Int16.Parse(cb_RegOcorrencia_Criticidade.Text.Substring(0, 2)) == 10)
                        impacto_cttu = "Alta";
                    #endregion

                    #region ABRE OCORRENCIA - INFORMAÇÕES

                    command = new MySqlCommand("INSERT INTO ocorrencia (cod_oco, cod_prj, cod_rqf, sistema, status_geral, identificador, atrib_para, classificacao, status_resposta, impacto_cttu) VALUES ('" +
                        COD_OCCORENCIA + "', '" +
                        tb_RegOcorrencia_Projeto.Text + "', '" +
                        tb_RegOcorrencia_RQF.Text + "', '" +
                        cb_RegOcorrencia_Sistema.Text +
                        "', 'Pendente FF', '" +
                        tb_RegOcorrencia_Identificador.Text +
                        "', '" + atribuidoPara + "', '" +
                        cb_RegOcorrencia_Classificacao.Text +
                        "', 'Aberta', '" +
                         impacto_cttu + "');", bdConn);
                    command.ExecuteNonQuery();

                    #endregion

                    #region CADASTRA DESCRIÇÃO DA OCORRÊNCIA

                    command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, descricao, questionamento, sugestao, tipo_registro, acao_analista) VALUES ('" +
                       COD_OCCORENCIA + "', '" +
                       tb_RegOcorrencia_Identificador.Text + "', '" +
                       tb_RegOcorrencia_Desc.Text + "', '" +
                       tb_RegOcorrencia_Quest.Text + "', '" +
                       tb_RegOcorrencia_Sugestao.Text + "', 'Abertura Ocorrencia', 'adicionou esta Abertura de Ocorrência');", bdConn);
                    command.ExecuteNonQuery();

                    #endregion

                    #region CADASTRA ANEXOS DA OCORRÊNCIA

                    if (countAnexo > 0)
                    {
                        pastaAnexoOco = pastaAnexo + @"\" + COD_OCCORENCIA;

                        if (!Directory.Exists(pastaAnexoOco))
                            Directory.CreateDirectory(pastaAnexoOco);
                    }

                    if (!String.IsNullOrEmpty(caminhoOrigemA1))
                    {
                        string pastaDestino = (pastaAnexoOco + @"\" + Path.GetFileName(caminhoOrigemA1));
                        File.Copy(caminhoOrigemA1, pastaDestino, true);

                        command = new MySqlCommand("INSERT INTO anexo_oco (cod_anexo, cod_oco, local_anexo) VALUES ('" +
                            COD_OCCORENCIA + "-A1', '" +
                            COD_OCCORENCIA + "', '" +
                            pastaDestino.Replace(@"\", @"\\") + "');", bdConn);
                        command.ExecuteNonQuery();
                    }

                    if (!String.IsNullOrEmpty(caminhoOrigemA2))
                    {
                        string pastaDestino = (pastaAnexoOco + @"\" + Path.GetFileName(caminhoOrigemA2));
                        File.Copy(caminhoOrigemA2, pastaDestino, true);

                        command = new MySqlCommand("INSERT INTO anexo_oco (cod_anexo, cod_oco, local_anexo) VALUES ('" +
                            COD_OCCORENCIA + "-A2', '" +
                            COD_OCCORENCIA + "', '" +
                             pastaDestino.Replace(@"\", @"\\") + "');", bdConn);
                        command.ExecuteNonQuery();
                    }

                    if (!String.IsNullOrEmpty(caminhoOrigemA3))
                    {
                        string pastaDestino = (pastaAnexoOco + @"\" + Path.GetFileName(caminhoOrigemA3));
                        File.Copy(caminhoOrigemA3, pastaDestino, true);

                        command = new MySqlCommand("INSERT INTO anexo_oco (cod_anexo, cod_oco, local_anexo) VALUES ('" +
                            COD_OCCORENCIA + "-A3', '" +
                            COD_OCCORENCIA + "', '" +
                             pastaDestino.Replace(@"\", @"\\") + "');", bdConn);
                        command.ExecuteNonQuery();
                    }

                    if (!String.IsNullOrEmpty(caminhoOrigemA4))
                    {
                        string pastaDestino = (pastaAnexoOco + @"\" + Path.GetFileName(caminhoOrigemA4));
                        File.Copy(caminhoOrigemA4, pastaDestino, true);

                        command = new MySqlCommand("INSERT INTO anexo_oco (cod_anexo, cod_oco, local_anexo) VALUES ('" +
                            COD_OCCORENCIA + "-A4', '" +
                            COD_OCCORENCIA + "', '" +
                             pastaDestino.Replace(@"\", @"\\") + "');", bdConn);
                        command.ExecuteNonQuery();
                    }

                    if (!String.IsNullOrEmpty(caminhoOrigemA5))
                    {
                        string pastaDestino = (pastaAnexoOco + @"\" + Path.GetFileName(caminhoOrigemA5));
                        File.Copy(caminhoOrigemA5, pastaDestino, true);

                        command = new MySqlCommand("INSERT INTO anexo_oco (cod_anexo, cod_oco, local_anexo) VALUES ('" +
                            COD_OCCORENCIA + "-A5', '" +
                            COD_OCCORENCIA + "', '" +
                             pastaDestino.Replace(@"\", @"\\") + "');", bdConn);
                        command.ExecuteNonQuery();
                    }

                    #endregion

                    //FECHA CONEXÃO
                    bdConn.Close();

                    MessageBox.Show("Ocorrência foi aberta com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    inicioRegistroOcorrencia();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }
        }

        //BOTÃO RETURN - PANEL 3 - REGISTOR DE OCORRENCIA
        private void bt_RegOcorrencia3_Return_Click(object sender, EventArgs e)
        {
            panel_RegOcorrencia_2.BringToFront();

            panel_RegOcorrencia_2.Visible = true;
            panel_RegOcorrencia_3.Visible = false;
        }

        //CONTROLE TEXTBOX DESCRIÇÃO - REGISTOR DE OCORRENCIA
        private void tb_RegOcorrencia_Desc_TextChanged(object sender, EventArgs e)
        {
            lb_RegOcorrencia_CrtlDesc.Text = (1000 - tb_RegOcorrencia_Desc.TextLength).ToString();
        }

        //CONTROLE TEXTBOX QUESTIONAMENTO - REGISTOR DE OCORRENCIA
        private void tb_RegOcorrencia_Quest_TextChanged(object sender, EventArgs e)
        {
            lb_RegOcorrencia_CrtlQuest.Text = (500 - tb_RegOcorrencia_Quest.TextLength).ToString();
        }

        //CONTROLE TEXTBOX SUGESTÃO/OBSERVAÇÃO - REGISTOR DE OCORRENCIA
        private void tb_RegOcorrencia_Sugestao_TextChanged(object sender, EventArgs e)
        {
            lb_RegOcorrencia_CrtlSug.Text = (500 - tb_RegOcorrencia_Sugestao.TextLength).ToString();
        }

        //VERIFICA CAMPOS - REGISTOR DE OCORRENCIA
        bool verificaCampos_RegOcorrencia()
        {
            if (tb_RegOcorrencia_Desc.Text == "")
            {
                MessageBox.Show("Escreva a Descrição da ocorrência.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            else
            {
                if (regexText(tb_RegOcorrencia_Desc.Text) == true)
                {
                    MessageBox.Show("O texto do campo (Descrição da Ocorrência) contém caracteres inválidos, favor verificar.\n\nCaracteres não permitidos: Aspas Simples e Contra-Barra", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            if (cb_RegOcorrencia_Classificacao.Text == "Dúvida")
                if (tb_RegOcorrencia_Quest.Text == "")
                {
                    MessageBox.Show("Escreva o Questionamento da ocorrência.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }

            if (regexText(tb_RegOcorrencia_Quest.Text) == true)
            {
                MessageBox.Show("O texto do campo (Questionamento) contém caracteres inválidos, favor verificar.\n\nCaracteres não permitidos: Aspas Simples e Contra-Barra", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (regexText(tb_RegOcorrencia_Sugestao.Text) == true && tb_RegOcorrencia_Sugestao.Text != "")
            {
                MessageBox.Show("O texto do campo (Sugestão/Observação) contém caracteres inválidos, favor verificar.\n\nCaracteres não permitidos: Aspas Simples e Contra-Barra", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        //REGEX
        public bool regexText(string input)
        {
            Regex nonTextRegex = new Regex(@"['\\]");
            if (nonTextRegex.IsMatch(input))
                return true;
            return false;
        }

        #region ANEXOS

        //BOTÃO ADICIONAR ANEXO
        private void bt_Anexo_Click(object sender, EventArgs e)
        {
            countAnexo++;

            if (countAnexo <= 5)
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    switch (countAnexo)
                    {
                        case 1:
                            caminhoOrigemA1 = openFileDialog1.FileName.ToString();
                            link_Anexo1.Text = Path.GetFileName(caminhoOrigemA1);
                            link_Anexo1.Visible = true;
                            lb_Anexos.Visible = true;
                            break;
                        case 2:
                            caminhoOrigemA2 = openFileDialog1.FileName.ToString();
                            link_Anexo2.Text = Path.GetFileName(caminhoOrigemA2);
                            link_Anexo2.Visible = true;
                            break;
                        case 3:
                            caminhoOrigemA3 = openFileDialog1.FileName.ToString();
                            link_Anexo3.Text = Path.GetFileName(caminhoOrigemA3);
                            link_Anexo3.Visible = true;
                            break;
                        case 4:
                            caminhoOrigemA4 = openFileDialog1.FileName.ToString();
                            link_Anexo4.Text = Path.GetFileName(caminhoOrigemA4);
                            link_Anexo4.Visible = true;
                            break;
                        case 5:
                            caminhoOrigemA5 = openFileDialog1.FileName.ToString();
                            link_Anexo5.Text = Path.GetFileName(caminhoOrigemA5);
                            link_Anexo5.Visible = true;
                            break;
                    }
                }
                else
                    countAnexo--;
            }
            else
                MessageBox.Show("Limite máximo de Anexo já foi adicionado!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //LINK ANEXO 1
        private void link_Anexo1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            form_Anexo anexo = new form_Anexo(caminhoOrigemA1);
            anexo.ShowDialog();

            if (remAnexo == 1)
            {
                link_Anexo1.Visible = false;
                caminhoOrigemA1 = "";
                countAnexo--;
                remAnexo = 0;

                if (countAnexo == 0)
                    lb_Anexos.Visible = false;
            }
        }

        //LINK ANEXO 2
        private void link_Anexo2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            form_Anexo anexo = new form_Anexo(caminhoOrigemA2);
            anexo.ShowDialog();

            if (remAnexo == 1)
            {
                link_Anexo2.Visible = false;
                caminhoOrigemA2 = "";
                countAnexo--;
                remAnexo = 0;

                if (countAnexo == 0)
                    lb_Anexos.Visible = false;
            }
        }

        //LINK ANEXO 3
        private void link_Anexo3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            form_Anexo anexo = new form_Anexo(caminhoOrigemA3);
            anexo.ShowDialog();

            if (remAnexo == 1)
            {
                link_Anexo3.Visible = false;
                caminhoOrigemA3 = "";
                countAnexo--;
                remAnexo = 0;

                if (countAnexo == 0)
                    lb_Anexos.Visible = false;
            }
        }

        //LINK ANEXO 4
        private void link_Anexo4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            form_Anexo anexo = new form_Anexo(caminhoOrigemA4);
            anexo.ShowDialog();

            if (remAnexo == 1)
            {
                link_Anexo4.Visible = false;
                caminhoOrigemA4 = "";
                countAnexo--;
                remAnexo = 0;

                if (countAnexo == 0)
                    lb_Anexos.Visible = false;
            }
        }

        //LINK ANEXO 5
        private void link_Anexo5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            form_Anexo anexo = new form_Anexo(caminhoOrigemA5);
            anexo.ShowDialog();

            if (remAnexo == 1)
            {
                link_Anexo5.Visible = false;
                caminhoOrigemA5 = "";
                countAnexo--;
                remAnexo = 0;

                if (countAnexo == 0)
                    lb_Anexos.Visible = false;
            }
        }

        #endregion

        #endregion

        #endregion

        #region //****************************************** MENU OCORRÊNCIAS  ******************************************\\

        //INICIO MENU OCORRENCIA
        void inicioMenuOcorrencias()
        {
            panel_MO_MinhasOcorrencias.Dock = DockStyle.Fill;
            panel_OcorrenciaDetalhes.Dock = DockStyle.Fill;
            panel_MO_BuscarProjetos1.Dock = DockStyle.Fill;
            panel_MO_BuscarProjetos2.Dock = DockStyle.Fill;
            panel_MO_OID.Dock = DockStyle.Fill;

            panel_MO_MinhasOcorrencias.SendToBack();
            panel_OcorrenciaDetalhes.SendToBack();
            panel_MO_BuscarProjetos1.SendToBack();
            panel_MO_BuscarProjetos2.SendToBack();
            panel_MO_OID.SendToBack();

            panel_MO_MinhasOcorrencias.Visible = false;
            panel_OcorrenciaDetalhes.Visible = false;
            panel_MO_BuscarProjetos1.Visible = false;
            panel_MO_BuscarProjetos2.Visible = false;
            panel_MO_OID.Visible = false;

        }

        //BOTÃO RETURN - MENU OCORRENCIAS
        private void bt_MO_Return_Click(object sender, EventArgs e)
        {
            panel_MO_MinhasOcorrencias.SendToBack();
            panel_OcorrenciaDetalhes.SendToBack();
            panel_MO_BuscarProjetos1.SendToBack();

            panel_MO_MinhasOcorrencias.Visible = false;
            panel_OcorrenciaDetalhes.Visible = false;
            panel_MO_BuscarProjetos1.Visible = false;

            atualizaPainelOcorrencias();
        }

        //MINHAS OCORRÊNCIAS
        private void bt_MO_MinhasOcorrencias_Click_1(object sender, EventArgs e)
        {
            panel_MO_MinhasOcorrencias.BringToFront();

            panel_MO_MinhasOcorrencias.Visible = true;

            inicioMinhasOcorrencias();
        }

        //BUSCA POR PROJETO
        private void bt_MO_BuscarPRJ_Click(object sender, EventArgs e)
        {
            panel_MO_BuscarProjetos1.BringToFront();

            panel_MO_BuscarProjetos1.Visible = true;
            panel_MO_BuscarProjetos2.Visible = false;

            tb_OP_SelectPRJ.Text = "";
        }

        //BUSCA POR ID
        private void bt_MO_BuscarID_Click(object sender, EventArgs e)
        {
            panel_MO_OID.BringToFront();

            panel_MO_OID.Visible = true;

            tb_OID_SelectID.Text = "P_____-OC___";
            tb_OID_SelectID.ForeColor = Color.Gray;
        }

        #region //****************************************** DETALHAMENTO OCORRÊNCIAS  ******************************************\\

        //ABRE - DETALHAMENTO DA OCORRENCIA
        void abreDetalhamentoOcorrencia(string CodigoOcorrencia, string retorno)
        {
            //ATIVA PANEL - DETALHAMENTO DA OCORRENCIA
            panel_OcorrenciaDetalhes.BringToFront();
            panel_OcorrenciaDetalhes.Visible = true;

            //CONTROLE GROUPBOX - DETALHAMENTO DA OCORRENCIA
            gb_DO_Detalhes.Visible = true;
            gb_DO_AdicionarObs.Visible = false;
            gb_DO_Resposta.Visible = false;
            DateTime data;

            //CONFIGURA RETORNO
            this.configVoltar = retorno;

            //RECUPERA INFORMAÇÕES DA OCORRENCIA
            try
            {
                controle_ComboboxDO = false;

                //ABRE CONEXÃO
                bdConn.Open();

                MySqlCommand command = new MySqlCommand("SELECT * FROM ocorrencia WHERE cod_oco = '" + CodigoOcorrencia + "';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                if (dr.Read())
                {
                    tb_DO_CodOcorrencia.Text = CodigoOcorrencia;
                    tb_DO_PRJ.Text = dr["cod_prj"].ToString();
                    tb_DO_RQF.Text = dr["cod_rqf"].ToString();
                    tb_DO_Sistema.Text = dr["sistema"].ToString();
                    tb_DO_Classificacao.Text = dr["classificacao"].ToString();
                    tb_DO_Identificador.Text = dr["identificador"].ToString();
                    tb_DO_AtribuidoPara.Text = dr["atrib_para"].ToString();
                    tb_DO_Criticidade.Text = dr["impacto_cttu"].ToString();
                    cb_ImpactoCTTU.Text = dr["impacto_cttu"].ToString();
                    tb_DO_StatusGeral.Text = dr["status_geral"].ToString();
                    tb_DO_StatusResposta.Text = dr["status_resposta"].ToString();
                    if (dr["dt_prev_solucao"].ToString() != "")
                    {
                        data = DateTime.Parse(dr["dt_prev_solucao"].ToString());
                        tb_DO_DataPrevistaSolucao.Text = String.Format("{0:dd/MM/yyyy}", data);
                    }


                }
                dr.Close();

                //VERIFICA QUAL EQUIPE É O USER
                if (UserON_Equipe == "Fábrica Desenvolvimento")
                {
                    #region USER FÁBRICA

                    //BOTÃO RESPONDER
                    bt_DO_Responder.Visible = false;

                    //BOTÃO ADD DATA PREVISTA                
                    bt_DO_AddDataPrev.Enabled = false;

                    //COMBOBOX ATRIBUIDO PARA
                    tb_DO_AtribuidoPara.Visible = true;
                    cb_DO_AtribuidoPara.Visible = false;

                    //COMBOBOX IMPACTO CTTU
                    if ((lb_UserON.Text == tb_DO_Identificador.Text) || (UserON_TipoUser == "Administrador"))
                    {
                        tb_DO_Criticidade.Visible = false;
                        cb_ImpactoCTTU.Visible = true;
                    }
                    else
                    {
                        tb_DO_Criticidade.Visible = true;
                        cb_ImpactoCTTU.Visible = false;
                    }

                    #endregion
                }
                else
                {
                    #region USER FUNCIONAL

                    //BOTÃO RESPONDER
                    bt_DO_Responder.Visible = true;

                    //BOTÃO ADD DATA PREVISTA
                    bt_DO_AddDataPrev.Enabled = true;

                    //COMBOBOX IMPACTO CTTU
                    tb_DO_Criticidade.Visible = true;
                    cb_ImpactoCTTU.Visible = false;

                    //COMBOBOX ATRIBUIDO PARA
                    if (UserON_TipoUser == "Administrador")
                    {
                        tb_DO_AtribuidoPara.Visible = false;
                        cb_DO_AtribuidoPara.Visible = true;

                        #region POVOA COMBOBOX ATRIBUÍDO PARA

                        cb_DO_AtribuidoPara.Items.Clear();

                        command = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_equipe = 'Fábrica Funcional' ORDER BY user_nome;", bdConn);
                        dr = command.ExecuteReader();
                        while (dr.Read())
                            cb_DO_AtribuidoPara.Items.Add(dr["user_nome"].ToString());
                        dr.Close();

                        #endregion

                        cb_DO_AtribuidoPara.Text = tb_DO_AtribuidoPara.Text;
                    }
                    else
                    {
                        tb_DO_AtribuidoPara.Visible = true;
                        cb_DO_AtribuidoPara.Visible = false;
                    }

                    #endregion
                }

                switch (tb_DO_Criticidade.Text)
                {
                    case "Média":
                        tb_DO_Criticidade.BackColor = Color.Yellow;
                        break;
                    case "Alta":
                        tb_DO_Criticidade.BackColor = Color.IndianRed;
                        break;
                    default:
                        tb_DO_Criticidade.BackColor = Color.LemonChiffon;
                        break;
                }

                //VERIFICA SE OCORRENCIA TEM ANEXO
                #region ANEXO 

                link_AnexoD1.Text = "";
                link_AnexoD2.Text = "";
                link_AnexoD3.Text = "";
                link_AnexoD4.Text = "";
                link_AnexoD5.Text = "";
                link_AnexoD1.LinkVisited = false;
                link_AnexoD2.LinkVisited = false;
                link_AnexoD3.LinkVisited = false;
                link_AnexoD4.LinkVisited = false;
                link_AnexoD5.LinkVisited = false;

                command = new MySqlCommand("SELECT local_anexo FROM anexo_oco WHERE cod_oco = '" + CodigoOcorrencia + "';", bdConn);
                dr = command.ExecuteReader();
                while (dr.Read())
                {
                    lb_Anexos_Det.Visible = true;

                    string local_anexo = dr["local_anexo"].ToString();
                    if (File.Exists(local_anexo))
                    {
                        if (String.IsNullOrEmpty(link_AnexoD1.Text))
                        {
                            link_AnexoD1.Text = Path.GetFileName(local_anexo);
                            link_AnexoD1.Visible = true;
                            caminhoOrigemA1D = local_anexo;
                            continue;
                        }
                        else if (String.IsNullOrEmpty(link_AnexoD2.Text))
                        {
                            link_AnexoD2.Text = Path.GetFileName(local_anexo);
                            link_AnexoD2.Visible = true;
                            caminhoOrigemA2D = local_anexo;
                            continue;
                        }
                        else if (String.IsNullOrEmpty(link_AnexoD3.Text))
                        {
                            link_AnexoD3.Text = Path.GetFileName(local_anexo);
                            link_AnexoD3.Visible = true;
                            caminhoOrigemA3D = local_anexo;
                            continue;
                        }
                        else if (String.IsNullOrEmpty(link_AnexoD4.Text))
                        {
                            link_AnexoD4.Text = Path.GetFileName(local_anexo);
                            link_AnexoD4.Visible = true;
                            caminhoOrigemA4D = local_anexo;
                            continue;
                        }
                        else if (String.IsNullOrEmpty(link_AnexoD5.Text))
                        {
                            link_AnexoD5.Text = Path.GetFileName(local_anexo);
                            link_AnexoD5.Visible = true;
                            caminhoOrigemA5D = local_anexo;
                            continue;
                        }

                    }

                }
                dr.Close();

                #endregion

                //ESCREVE OS DETALHES DA OCORRENCIA
                command = new MySqlCommand("SELECT * FROM desc_ocorrencia WHERE cod_oco = '" + CodigoOcorrencia + "';", bdConn);
                dr = command.ExecuteReader();
                tb_DetalhesOcorrencia_Descrição.Clear();
                tb_DetalhesOcorrencia_DescriçãoAbertura.Clear();
                while (dr.Read())
                {
                    if (dr["tipo_registro"].ToString() == "Abertura Ocorrencia")
                    {
                        if (!String.IsNullOrEmpty(dr["descricao"].ToString()))
                        {
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionColor = Color.Red;
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 12, FontStyle.Bold);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText("Ocorrência:\n");
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 10);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText(dr["descricao"].ToString() + "\n\n");
                        }

                        if (!String.IsNullOrEmpty(dr["questionamento"].ToString()))
                        {
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionColor = Color.Red;
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 12, FontStyle.Bold);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText("Questionamento:\n");
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 10);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText(dr["questionamento"].ToString() + "\n\n");
                        }

                        if (!String.IsNullOrEmpty(dr["sugestao"].ToString()))
                        {
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionColor = Color.Red;
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 12, FontStyle.Bold);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText("Sugestão / Observação:\n");
                            tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 10);
                            tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText(dr["sugestao"].ToString() + "\n\n");
                        }

                        tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionColor = Color.Gray;
                        tb_DetalhesOcorrencia_DescriçãoAbertura.SelectionFont = new System.Drawing.Font("Arial", 8, FontStyle.Italic);
                        tb_DetalhesOcorrencia_DescriçãoAbertura.AppendText(dr["analista"].ToString() + " " + dr["acao_analista"].ToString() + ", em " + dr["horario_registro"].ToString() + "\n");

                    }
                    else
                    {

                        tb_DetalhesOcorrencia_Descrição.AppendText("----------------------------------------------------------------------------------------\n");

                        if (!String.IsNullOrEmpty(dr["descricao"].ToString()))
                        {
                            tb_DetalhesOcorrencia_Descrição.SelectionColor = Color.Red;
                            tb_DetalhesOcorrencia_Descrição.SelectionFont = new System.Drawing.Font("Arial", 12, FontStyle.Bold);
                            tb_DetalhesOcorrencia_Descrição.AppendText((dr["tipo_registro"].ToString() == "Resposta") ? "Resposta:\n" : "Descrição:\n");
                            tb_DetalhesOcorrencia_Descrição.SelectionFont = new System.Drawing.Font("Arial", 10);
                            tb_DetalhesOcorrencia_Descrição.AppendText(dr["descricao"].ToString() + "\n\n");
                        }

                        if (!String.IsNullOrEmpty(dr["sugestao"].ToString()))
                        {
                            tb_DetalhesOcorrencia_Descrição.SelectionColor = Color.Red;
                            tb_DetalhesOcorrencia_Descrição.SelectionFont = new System.Drawing.Font("Arial", 12, FontStyle.Bold);
                            tb_DetalhesOcorrencia_Descrição.AppendText("Sugestão / Observação:\n");
                            tb_DetalhesOcorrencia_Descrição.SelectionFont = new System.Drawing.Font("Arial", 10);
                            tb_DetalhesOcorrencia_Descrição.AppendText(dr["sugestao"].ToString() + "\n\n");
                        }


                        tb_DetalhesOcorrencia_Descrição.SelectionColor = Color.Gray;
                        tb_DetalhesOcorrencia_Descrição.SelectionFont = new System.Drawing.Font("Arial", 8, FontStyle.Italic);
                        tb_DetalhesOcorrencia_Descrição.AppendText(dr["analista"].ToString() + " " + dr["acao_analista"].ToString() + ", em " + dr["horario_registro"].ToString() + "\n");

                    }
                }
                dr.Close();

                //FECHA CONEXÃO
                bdConn.Close();

                //TRAVA COMBOBOX SELECT CHANGE
                controle_ComboboxDO = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO ADICIONAR OBSERVAÇÃO - DETALHAMENTO DA OCORRENCIA
        private void bt_DO_AdicionarObs_Click(object sender, EventArgs e)
        {
            gb_DO_AdicionarObs.BringToFront();
            gb_DO_AdicionarObs.Visible = true;
            gb_DO_Detalhes.Visible = false;
            gb_DO_Resposta.Visible = false;

            tb_DO_AddObs.Text = "";
        }

        //BOTÃO SALVAR - DETALHAMENTO DA OCORRENCIA
        private void bt_DO_SalvarStatus_Click(object sender, EventArgs e)
        {
            try
            {
                if (UserON_Equipe == "Fábrica Desenvolvimento")
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("UPDATE ocorrencia SET status_geral = '" + tb_DO_StatusGeral.Text + "' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                    command.ExecuteNonQuery();

                    //FECHA CONEXÃO
                    bdConn.Close();

                    MessageBox.Show("Status atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("UPDATE ocorrencia SET status_resposta = '" + tb_DO_StatusResposta.Text + "' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                    command.ExecuteNonQuery();

                    //FECHA CONEXÃO
                    bdConn.Close();

                    MessageBox.Show("Status atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO EXPOTAR OCORRENCIA - EXPOTAR OCORRENCIA
        private void bt_ExportarOcorrencia_Click(object sender, EventArgs e)
        {
            try
            {
                //CRIA OBJETO WORD
                Word.Application wordApp = new Word.Application();
                string diretorio = @"C:\ControleOcorrencia";
                Word.Document wordDoc = wordApp.Documents.Add(diretorio + @"\Template_CrtlOco.docx");

                #region REPLACE NO DOCUMENTO

                Version versao = Assembly.GetExecutingAssembly().GetName().Version;
                ReplaceBookmarkText(wordDoc, "bkVersao", (versao.ToString().Substring(0, 3)));

                if (tb_DO_CodOcorrencia.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkCodOco", tb_DO_CodOcorrencia.Text);

                if (tb_DO_PRJ.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkProjeto", tb_DO_PRJ.Text);

                if (tb_DO_RQF.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkRQF", tb_DO_RQF.Text);

                if (tb_DO_Sistema.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkSistema", tb_DO_Sistema.Text);

                if (tb_DO_Classificacao.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkClassificacao", tb_DO_Classificacao.Text);

                if (tb_DO_Identificador.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkIdentificador", tb_DO_Identificador.Text);

                if (tb_DO_AtribuidoPara.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkAtribuidoP", tb_DO_AtribuidoPara.Text);

                if (cb_DO_Impacto.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkImpactoCTTU", cb_DO_Impacto.Text);

                if (tb_DO_Criticidade.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkImpactoCTTU", tb_DO_Criticidade.Text);

                if (tb_DO_DataPrevistaSolucao.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkDtPrevista", tb_DO_DataPrevistaSolucao.Text);
                else
                    ReplaceBookmarkText(wordDoc, "bkDtPrevista", "-");

                if (tb_DO_StatusGeral.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkStatusGeral", tb_DO_StatusGeral.Text);

                if (tb_DO_StatusResposta.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkStatusResposta", tb_DO_StatusResposta.Text);

                if (tb_DetalhesOcorrencia_DescriçãoAbertura.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkDescricaoOco", tb_DetalhesOcorrencia_DescriçãoAbertura.Text);

                if (tb_DetalhesOcorrencia_Descrição.Text != "")
                    ReplaceBookmarkText(wordDoc, "bkDetalhamentoOco", tb_DetalhesOcorrencia_Descrição.Text.ToString());

                #endregion

                //OBJETO PARA SALVAR
                SaveFileDialog salvar = new SaveFileDialog();
                salvar.Title = "Exportar Ocorrência";
                salvar.Filter = "Arquivo do Word *.docx | *.docx";
                salvar.FileName = "Descrição Ocorrencia - " + tb_DO_CodOcorrencia.Text;
                DialogResult result = salvar.ShowDialog();

                if (result == DialogResult.OK)
                {
                    wordDoc.SaveAs(salvar.FileName);
                    wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    wordApp.Quit();
                    MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                    wordApp.Quit(false);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        //ALTERAÇÃO DE TEXTO NO DOCUMENTO
        private void ReplaceBookmarkText(Word.Document doc, string bookmarkName, string text)
        {
            if (doc.Bookmarks.Exists(bookmarkName))
            {
                Object name = bookmarkName;
                Word.Range range = doc.Bookmarks.get_Item(ref name).Range;
                range.Text = text;
                object newRange = range;
                doc.Bookmarks.Add(bookmarkName, ref newRange);
            }
        }

        //BOTÃO ADICIONAR RESPOSTA - DETALHAMENTO DA OCORRENCIA
        private void bt_DO_Responder_Click(object sender, EventArgs e)
        {
            bool encerraOco = false;

            bdConn.Open();
            MySqlCommand command = new MySqlCommand("SELECT encerra_ocorrencia FROM status_geral WHERE nome_status = '" + tb_DO_StatusGeral.Text + "';", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            if (dr.Read())
                encerraOco = ((dr["encerra_ocorrencia"].ToString() == "sim") ? true : false);
            dr.Close();
            bdConn.Close();

            if (!encerraOco)
            {
                gb_DO_Resposta.BringToFront();
                gb_DO_Resposta.Visible = true;
                gb_DO_Detalhes.Visible = false;
                gb_DO_AdicionarObs.Visible = false;

                tb_DO_Resp_Resposta.Text = "";

                try
                {
                    cb_DO_Impacto.Items.Clear();

                    //ABRE CONEXÃO
                    bdConn.Open();

                    command = new MySqlCommand("SELECT * FROM status_impacto;", bdConn);
                    dr = command.ExecuteReader();

                    while (dr.Read())
                        cb_DO_Impacto.Items.Add(dr["nome_status"].ToString());
                    dr.Close();

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }
            else
                MessageBox.Show("Não é possivel modificar está ocorrência.", "Ops...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //BOTÃO HOME - DETALHAMENTO DA OCORRENCIA
        private void bt_DO_Home_Click(object sender, EventArgs e)
        {
            inicioMenuOcorrencias();
        }

        //ABRE FORM PARA ATUALIZAR STATUS GERAL
        private void bt_DO_AtualizaStatusGeral_Click(object sender, EventArgs e)
        {
            bool encerraOco = false;

            bdConn.Open();
            MySqlCommand command = new MySqlCommand("SELECT encerra_ocorrencia FROM status_geral WHERE nome_status = '" + tb_DO_StatusGeral.Text + "';", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            if (dr.Read())
                encerraOco = ((dr["encerra_ocorrencia"].ToString() == "sim") ? true : false);
            dr.Close();
            bdConn.Close();

            if (!encerraOco)
            {
                form_StatusGeral fStatusGeral = new form_StatusGeral(tb_DO_CodOcorrencia.Text, lb_UserON.Text);
                fStatusGeral.ShowDialog();

                abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
            }
            else
                MessageBox.Show("Não é possivel modificar está ocorrência.", "Ops...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //ABRE FORM PARA ATUALIZAR STATUS RESPOSTA
        private void bt_DO_AtualizaStatusResposta_Click(object sender, EventArgs e)
        {
            bool encerraOco = false;

            bdConn.Open();
            MySqlCommand command = new MySqlCommand("SELECT encerra_ocorrencia FROM status_geral WHERE nome_status = '" + tb_DO_StatusGeral.Text + "';", bdConn);
            MySqlDataReader dr = command.ExecuteReader();
            if (dr.Read())
                encerraOco = ((dr["encerra_ocorrencia"].ToString() == "sim") ? true : false);
            dr.Close();
            bdConn.Close();

            if (!encerraOco)
            {
                form_StatusResposta fStatusResposta = new form_StatusResposta(tb_DO_CodOcorrencia.Text, lb_UserON.Text, false);
                fStatusResposta.ShowDialog();

                abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
            }
            else
                MessageBox.Show("Não é possivel modificar está ocorrência.", "Ops...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        //ADICIONAR DATA PREVISTA SOLUÇÃO
        private void bt_DO_AddDataPrev_Click(object sender, EventArgs e)
        {
            form_AddDataPrev fr_DataPrev = new form_AddDataPrev(tb_DO_CodOcorrencia.Text, lb_UserON.Text);
            fr_DataPrev.ShowDialog();

            abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
        }

        //COMBOBOX ALTERAR ATRIBUIDO PARA
        private void cb_DO_AtribuidoPara_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (controle_ComboboxDO == true)
                {
                    string atribPara_Atual = "";

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT atrib_para FROM ocorrencia WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    if (dr.Read())
                        atribPara_Atual = dr["atrib_para"].ToString();
                    dr.Close();

                    if (cb_DO_AtribuidoPara.Text != atribPara_Atual)
                    {
                        DialogResult result = MessageBox.Show("Deseja atríbuir essa ocorrência para " + cb_DO_AtribuidoPara.Text + "?", "Atribuído Para", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            string email_de = "";
                            string email_para = "";

                            //ATUALIZA OCORRENCIA - ATRIBUI PARA NOVO ANALISTA
                            command = new MySqlCommand("UPDATE ocorrencia SET atrib_para = '" + cb_DO_AtribuidoPara.Text + "' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                            command.ExecuteNonQuery();

                            //ADICIONA DESCRIÇÃO (LOG)
                            command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + tb_DO_CodOcorrencia.Text + "','" + lb_UserON.Text + "', 'Atribuir Para', 'atribuiu está ocorrência para " + cb_DO_AtribuidoPara.Text + "');", bdConn);
                            command.ExecuteNonQuery();

                            //RECUPERA EMAIL DE (USUARIO LOGADO)
                            command = new MySqlCommand("SELECT user_email FROM usuarios WHERE user_nome = '" + lb_UserON.Text + "';", bdConn);
                            dr = command.ExecuteReader();
                            if (dr.Read())
                                email_de = dr["user_email"].ToString();
                            dr.Close();

                            //RECUPERA EMAIL PARA (USUARIO A QUEM ESTÁ SENDO ATRIBUIDO)
                            command = new MySqlCommand("SELECT user_email FROM usuarios WHERE user_nome = '" + cb_DO_AtribuidoPara.Text + "';", bdConn);
                            dr = command.ExecuteReader();
                            if (dr.Read())
                                email_para = dr["user_email"].ToString();
                            dr.Close();

                            //ALTERAR ANTES DE VERSIONAR
                            //victor.hugo.o.santos@accenture.com
                            //CRIA OBJETO DO EMAIL
                            /*Outlook.Application oApp = new Outlook.Application();
                            SendEmailFromAccount(oApp,
                                "Controle de Ocorrência - " + tb_DO_CodOcorrencia.Text + "",
                                criaEmailBody(cb_DO_AtribuidoPara.Text, tb_DO_CodOcorrencia.Text, tb_DO_PRJ.Text, tb_DO_RQF.Text),
                                "victor.hugo.o.santos@accenture.com",
                                "victor.hugo.o.santos@accenture.com");*/

                            Outlook.Application oApp = new Outlook.Application();
                            SendEmailFromAccount(oApp,
                                "Controle de Ocorrência - " + tb_DO_CodOcorrencia.Text + "",
                                criaEmailBody(cb_DO_AtribuidoPara.Text, tb_DO_CodOcorrencia.Text, tb_DO_PRJ.Text, tb_DO_RQF.Text),
                                email_para,
                                email_de);
                        }
                        else
                        {
                            //FECHA CONEXÃO
                            bdConn.Close();

                            cb_DO_AtribuidoPara.Text = atribPara_Atual;
                        }

                    }

                    //FECHA CONEXÃO
                    bdConn.Close();

                    abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro ao alterar 'Atribuido Para'", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //COMBOBOX ALTERAR IMPACTO CTTU
        private void cb_ImpactoCTTU_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (controle_ComboboxDO == true)
            {
                try
                {
                    string impacto_Atual = "";

                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command = new MySqlCommand("SELECT impacto_cttu FROM ocorrencia WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                    MySqlDataReader dr = command.ExecuteReader();
                    if (dr.Read())
                        impacto_Atual = dr["impacto_cttu"].ToString();
                    dr.Close();

                    if (cb_ImpactoCTTU.Text != impacto_Atual)
                    {
                        DialogResult result = MessageBox.Show("Deseja alterar o Impacto CTTU para " + cb_ImpactoCTTU.Text + "?", "Impacto CTTU", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                        if (result == DialogResult.Yes)
                        {
                            //ATUALIZA OCORRENCIA - ATRIBUI PARA NOVO ANALISTA
                            command = new MySqlCommand("UPDATE ocorrencia SET impacto_cttu = '" + cb_ImpactoCTTU.Text + "' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                            command.ExecuteNonQuery();

                            //ADICIONA DESCRIÇÃO (LOG)
                            command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + tb_DO_CodOcorrencia.Text + "','" + lb_UserON.Text + "', 'Alteração Impacto CTTU', 'alterou o Impacto CTTU para " + cb_ImpactoCTTU.Text + "');", bdConn);
                            command.ExecuteNonQuery();
                        }
                    }
                    else
                        cb_ImpactoCTTU.Text = impacto_Atual;

                    //FECHA CONEXÃO
                    bdConn.Close();

                    abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro ao alterar 'Impacto CTTU'", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        //BOTÃO VOLTAR
        private void bt_Voltar_Click(object sender, EventArgs e)
        {
            switch (configVoltar)
            {
                case "Minhas Ocorrencia":
                    minhasOcorrêciaToolStripMenuItem.PerformClick();
                    break;
                case "Por ID":
                    porIDToolStripMenuItem.PerformClick();
                    break;
                case "Por Projeto":
                    porProjetoToolStripMenuItem.PerformClick();
                    abrePesquisaPRJ(tb_DO_PRJ.Text);
                    lb_OP_PRJFiltros.Text = tb_DO_PRJ.Text;
                    panel_MO_BuscarProjetos2.BringToFront();
                    panel_MO_BuscarProjetos2.Dock = DockStyle.Fill;
                    panel_MO_BuscarProjetos2.Visible = true;
                    break;
                default:
                    home();
                    break;
            }
        }

        #region ADICIONAR OBSERVAÇÃO

        //BOTÃO SALVAR - ADICIONAR OBSERVAÇÃO
        private void bt_DO_AddObs_Salvar_Click(object sender, EventArgs e)
        {
            try
            {
                if (!String.IsNullOrEmpty(tb_DO_AddObs.Text))
                    if (regexText(tb_DO_AddObs.Text) == false)
                    {
                        //ABRE CONEXÃO
                        bdConn.Open();

                        MySqlCommand command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, sugestao, tipo_registro, acao_analista) VALUES ('" + tb_DO_CodOcorrencia.Text + "','" + lb_UserON.Text + "','" + tb_DO_AddObs.Text + "', 'Observacao', 'adicionou esta Observação');", bdConn);
                        command.ExecuteNonQuery();

                        //FECHA CONEXÃO
                        bdConn.Close();

                        MessageBox.Show("Observação foi adicionada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
                    }
                    else
                        MessageBox.Show("O texto do campo (Observação) contém caracteres inválidos, favor verificar.\n\nCaracteres não permitidos: Aspas Simples e Contra-Barra", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                else
                    MessageBox.Show("Preencha o campo de observação para salvar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO LIMPAR - ADICIONAR OBSERVAÇÃO
        private void bt_DO_AddObs_Limpar_Click(object sender, EventArgs e)
        {
            tb_DO_AddObs.Text = "";
        }

        //BOTÃO CANCELAR - ADICIONAR OBSERVAÇÃO
        private void bt_DO_AddObs_Cancelar_Click(object sender, EventArgs e)
        {
            gb_DO_Detalhes.BringToFront();
            gb_DO_Detalhes.Visible = true;
            gb_DO_AdicionarObs.Visible = false;
            gb_DO_Resposta.Visible = false;

            tb_DO_AddObs.Text = "";
        }

        //CONTROLE TEXTBOX OBSERVAÇÃO - ADICIONAR OBSERVAÇÃO
        private void tb_DO_AddObs_TextChanged(object sender, EventArgs e)
        {
            lb_DO_AddObs_CrtlObs.Text = (500 - tb_DO_AddObs.TextLength).ToString();
        }

        #endregion

        #region ADICIONAR RESPOSTA

        //BOTÃO SALVAR - ADICIONAR RESPOSTA
        private void bt_DO_Resp_Salvar_Click(object sender, EventArgs e)
        {

            try
            {
                if (validaCamposAddResposta())
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    MySqlCommand command;

                    command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, descricao, tipo_registro, acao_analista) VALUES ('" + tb_DO_CodOcorrencia.Text + "','" + lb_UserON.Text + "','" + tb_DO_Resp_Resposta.Text + "', 'Resposta', 'adicionou esta Resposta');", bdConn);
                    command.ExecuteNonQuery();

                    if (rb_DO_ImpactoSim.Checked == true)
                    {
                        command = new MySqlCommand("SELECT email_testes FROM status_impacto WHERE nome_status = '" + cb_DO_Impacto.Text + "';", bdConn);
                        MySqlDataReader dr = command.ExecuteReader();
                        if (dr.Read())
                            if (dr["email_testes"].ToString() == "Sim")
                                enviaRelatorio_Testes(tb_DO_PRJ.Text);
                        dr.Close();


                        command = new MySqlCommand("UPDATE ocorrencia SET status_geral = '" + cb_DO_Impacto.Text + "' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                        command.ExecuteNonQuery();

                        command = new MySqlCommand("UPDATE ocorrencia SET status_resposta = 'Respondido' WHERE cod_oco = '" + tb_DO_CodOcorrencia.Text + "';", bdConn);
                        command.ExecuteNonQuery();

                        command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + tb_DO_CodOcorrencia.Text + "', '" + lb_UserON.Text + "', 'Atualizacao Status', 'atualizou o Status Geral (Impacto) para " + cb_DO_Impacto.Text + "');", bdConn);
                        command.ExecuteNonQuery();

                        MessageBox.Show("Sua resposta foi adicionada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Sua resposta foi adicionada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        DialogResult result = MessageBox.Show("É necessário atualizar o Status de Resposta!\n\nDeseja atualizar agora?", "Atenção!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            form_StatusResposta frResposta = new form_StatusResposta(tb_DO_CodOcorrencia.Text, lb_UserON.Text, ((rb_DO_ImpactoSim.Checked) ? true : false));
                            frResposta.ShowDialog();
                        }
                    }

                    //FECHA CONEXÃO
                    bdConn.Close();

                    abreDetalhamentoOcorrencia(tb_DO_CodOcorrencia.Text, this.configVoltar);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro ao adicionar resposta!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        //VALIDA CAMPOS ADICIONAR RESPOSTA
        bool validaCamposAddResposta()
        {
            if (rb_DO_ImpactoSim.Checked == true)
                if (cb_DO_Impacto.Text == "")
                {
                    MessageBox.Show("Selecione um impacto!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }

            if (tb_DO_Resp_Resposta.Text == "")
            {
                MessageBox.Show("Preencha o campo de resposta para salvar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
            else
            {
                if (regexText(tb_DO_Resp_Resposta.Text) == true)
                {
                    MessageBox.Show("O texto do campo (Resposta) contém caracteres inválidos, favor verificar.\n\nCaracteres não permitidos: Aspas Simples e Contra-Barra", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return false;
                }
            }

            return true;
        }

        //BOTÃO LIMPAR - ADICIONAR RESPOSTA
        private void bt_DO_Resp_Limpar_Click(object sender, EventArgs e)
        {
            tb_DO_Resp_Resposta.Text = "";
            rb_DO_ImpactoNao.Checked = true;
            cb_DO_Impacto.Text = null;
        }

        //BOTÃO CANCELAR - ADICIONAR RESPOSTA
        private void bt_DO_Resp_Cancelar_Click(object sender, EventArgs e)
        {
            gb_DO_Detalhes.BringToFront();
            gb_DO_Detalhes.Visible = true;
            gb_DO_Resposta.Visible = false;
            gb_DO_AdicionarObs.Visible = false;

            tb_DO_Resp_Resposta.Text = "";
        }

        //CONTROLE TEXTBOX RESPOSTA - ADICIONAR RESPOSTA
        private void tb_DO_Resp_Resposta_TextChanged(object sender, EventArgs e)
        {
            lb_DO_Resp_CrtlResposta.Text = (1000 - tb_DO_Resp_Resposta.TextLength).ToString();
        }

        //VERIFICA SE EXISTE IMPACTO NA OCORRENCIA
        private void rb_DO_ImpactoNao_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_DO_ImpactoNao.Checked == false)
                gb_DO_Impacto.Visible = true;
            else
                gb_DO_Impacto.Visible = false;
        }

        #endregion

        #region ANEXOS

        //LINK ANEXO 1
        private void link_AnexoD1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                link_AnexoD1.LinkVisited = true;
                if (File.Exists(caminhoOrigemA1D))
                {
                    DialogResult result = MessageBox.Show("Deseja baixar o anexo?", "Salvar Anexo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        //OBJETO PARA SALVAR
                        SaveFileDialog salvar = new SaveFileDialog();
                        salvar.Title = "Salvar Anexo";
                        salvar.Filter = "Arquivo do Word *.docx | *.docx|Excel *.xlsx | *.xlsx|Arquivo de Texto *.txt | *.txt|PDF *.pdf | *.pdf|JPeg Image|*.jpg|PNG File *.png | *.png";
                        salvar.FileName = link_AnexoD1.Text;
                        result = salvar.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            File.Copy(caminhoOrigemA1D, salvar.FileName, true);
                            MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                    MessageBox.Show("Arquivo não existe no diretório indicado!\nLocal: " + caminhoOrigemA1D + "\n\nCaso o problema persista, contate o gerenciador!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arquivo indisponível no momento! Caso o problema persista, contate o gerenciador!\n\nDetalhes: " + ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //LINK ANEXO 2
        private void link_AnexoD2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                link_AnexoD2.LinkVisited = true;
                if (File.Exists(caminhoOrigemA2D))
                {
                    DialogResult result = MessageBox.Show("Deseja baixar o anexo?", "Salvar Anexo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        //OBJETO PARA SALVAR
                        SaveFileDialog salvar = new SaveFileDialog();
                        salvar.Title = "Salvar Anexo";
                        salvar.Filter = "Arquivo do Word *.docx | *.docx|Excel *.xlsx | *.xlsx|Arquivo de Texto *.txt | *.txt|PDF *.pdf | *.pdf|JPeg Image|*.jpg|PNG File *.png | *.png";
                        salvar.FileName = link_AnexoD2.Text;
                        result = salvar.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            File.Copy(caminhoOrigemA2D, salvar.FileName, true);
                            MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                    MessageBox.Show("Arquivo não existe no diretório indicado!\nLocal: " + caminhoOrigemA2D + "\n\nCaso o problema persista, contate o gerenciador!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arquivo indisponível no momento! Caso o problema persista, contate o gerenciador!\n\nDetalhes: " + ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //LINK ANEXO 3
        private void link_AnexoD3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                link_AnexoD3.LinkVisited = true;
                if (File.Exists(caminhoOrigemA3D))
                {
                    DialogResult result = MessageBox.Show("Deseja baixar o anexo?", "Salvar Anexo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        //OBJETO PARA SALVAR
                        SaveFileDialog salvar = new SaveFileDialog();
                        salvar.Title = "Salvar Anexo";
                        salvar.Filter = "Arquivo do Word *.docx | *.docx|Excel *.xlsx | *.xlsx|Arquivo de Texto *.txt | *.txt|PDF *.pdf | *.pdf|JPeg Image|*.jpg|PNG File *.png | *.png";
                        salvar.FileName = link_AnexoD3.Text;
                        result = salvar.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            File.Copy(caminhoOrigemA3D, salvar.FileName, true);
                            MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                    MessageBox.Show("Arquivo não existe no diretório indicado!\nLocal: " + caminhoOrigemA3D + "\n\nCaso o problema persista, contate o gerenciador!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arquivo indisponível no momento! Caso o problema persista, contate o gerenciador!\n\nDetalhes: " + ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //LINK ANEXO 4
        private void link_AnexoD4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                link_AnexoD4.LinkVisited = true;
                if (File.Exists(caminhoOrigemA4D))
                {
                    DialogResult result = MessageBox.Show("Deseja baixar o anexo?", "Salvar Anexo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        //OBJETO PARA SALVAR
                        SaveFileDialog salvar = new SaveFileDialog();
                        salvar.Title = "Salvar Anexo";
                        salvar.Filter = "Arquivo do Word *.docx | *.docx|Excel *.xlsx | *.xlsx|Arquivo de Texto *.txt | *.txt|PDF *.pdf | *.pdf|JPeg Image|*.jpg|PNG File *.png | *.png";
                        salvar.FileName = link_AnexoD4.Text;
                        result = salvar.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            File.Copy(caminhoOrigemA4D, salvar.FileName, true);
                            MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                    MessageBox.Show("Arquivo não existe no diretório indicado!\nLocal: " + caminhoOrigemA4D + "\n\nCaso o problema persista, contate o gerenciador!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arquivo indisponível no momento! Caso o problema persista, contate o gerenciador!\n\nDetalhes: " + ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        //LINK ANEXO 5
        private void link_AnexoD5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                link_AnexoD5.LinkVisited = true;
                if (File.Exists(caminhoOrigemA5D))
                {
                    DialogResult result = MessageBox.Show("Deseja baixar o anexo?", "Salvar Anexo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        //OBJETO PARA SALVAR
                        SaveFileDialog salvar = new SaveFileDialog();
                        salvar.Title = "Salvar Anexo";
                        salvar.Filter = "Arquivo do Word *.docx | *.docx|Excel *.xlsx | *.xlsx|Arquivo de Texto *.txt | *.txt|PDF *.pdf | *.pdf|JPeg Image|*.jpg|PNG File *.png | *.png";
                        salvar.FileName = link_AnexoD5.Text;
                        result = salvar.ShowDialog();

                        if (result == DialogResult.OK)
                        {
                            File.Copy(caminhoOrigemA5D, salvar.FileName, true);
                            MessageBox.Show("Documento gerado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                else
                    MessageBox.Show("Arquivo não existe no diretório indicado!\nLocal: " + caminhoOrigemA5D + "\n\nCaso o problema persista, contate o gerenciador!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Arquivo indisponível no momento! Caso o problema persista, contate o gerenciador!\n\nDetalhes: " + ex.Message, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        #endregion

        #region EMAIL

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
        static string criaEmailBody(string responsavel, string ocorrencia, string projeto, string rqf)
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

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Uma ocorrência foi atríbuida para você.<o:p></o:p></span></p>  ";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><u>Detalhes:</u><o:p></o:p></span></p>  ";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>Projeto: </b>" + projeto + "<o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>RQF: </b>" + rqf + "<o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'><b>Código Ocorrência: </b><span style='color:red'>" + ocorrencia + "</span><o:p></o:p></span></p>";
            emailBody += "  <p class=MsoNormal><span style='font-family:'Arial',sans-serif'><o:p>&nbsp;</o:p></span></p>";

            emailBody += "  <p class=MsoNormal><span style='font-size:12.0pt'>Favor acessar a ferramenta “<i>Controle de Ocorrências</i>” para outras informações.<o:p></o:p></span></p>";
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

        //EMAIL - ENVIA RELATÓRIO EQUIPE TESTES
        void enviaRelatorio_Testes(string Projeto)
        {

        }

        #endregion

        #endregion

        #region //****************************************** MINHAS OCORRÊNCIAS  ******************************************\\

        //INICIO MENU OCORRENCIA
        void inicioMinhasOcorrencias()
        {
            //LIMPA GRIDVIEW
            if (this.dataGrid_MinhasOcorrencias.DataSource != null)
                this.dataGrid_MinhasOcorrencias.DataSource = null;
            else
            {
                this.dataGrid_MinhasOcorrencias.Rows.Clear();
                this.dataGrid_MinhasOcorrencias.Columns.Clear();
            }

            //HEADER DATAGRIDVIEW            
            dataGrid_MinhasOcorrencias.Columns.Add("cod_prj", "Projeto");
            dataGrid_MinhasOcorrencias.Columns.Add("dfrqf", "DF/RQF");
            dataGrid_MinhasOcorrencias.Columns.Add("sistema", "Sistema");
            dataGrid_MinhasOcorrencias.Columns.Add("cod_oco", "Cód. Ocorrência");
            dataGrid_MinhasOcorrencias.Columns.Add("identificador", "Identificador");
            dataGrid_MinhasOcorrencias.Columns.Add("classificacao", "Classificação");
            dataGrid_MinhasOcorrencias.Columns.Add("impacto_cttu", "Impacto CTTU");
            dataGrid_MinhasOcorrencias.Columns.Add("status_geral", "Status Geral");
            dataGrid_MinhasOcorrencias.Columns.Add("atrib_para", "Atribuído Para");
            dataGrid_MinhasOcorrencias.Columns.Add("status_resp", "Status da Resposta");
            dataGrid_MinhasOcorrencias.Columns.Add("dt_prev_sol", "DT Prevista Solução");

            //ABRE CONEXÃO
            bdConn.Open();

            MySqlCommand command;

            if (UserON_Equipe == "Fábrica Desenvolvimento")
                command = new MySqlCommand("SELECT * FROM ocorrencia WHERE identificador = '" + lb_UserON.Text + "' ORDER BY cod_prj, cod_rqf, cod_oco;", bdConn);
            else
                command = new MySqlCommand("SELECT * FROM ocorrencia WHERE atrib_para = '" + lb_UserON.Text + "' ORDER BY cod_prj, cod_rqf, cod_oco;", bdConn);

            MySqlDataReader dr = command.ExecuteReader();
            while (dr.Read())
            {
                dataGrid_MinhasOcorrencias.Rows.Add(
                dr["cod_prj"].ToString(),
                dr["cod_rqf"].ToString(),
                dr["sistema"].ToString(),
                dr["cod_oco"].ToString(),
                dr["identificador"].ToString(),
                dr["classificacao"].ToString(),
                dr["impacto_cttu"].ToString(),
                dr["status_geral"].ToString(),
                dr["atrib_para"].ToString(),
                dr["status_resposta"].ToString(),
                ((dr["dt_prev_solucao"].ToString() != "") ? String.Format("{0:dd/MM/yyyy}", DateTime.Parse(dr["dt_prev_solucao"].ToString())) : "-")
                );
            }
            //BOTÃO ABRIR
            DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
            dataGrid_MinhasOcorrencias.Columns.Add(btn);
            btn.HeaderText = "Alteração";
            btn.Text = "Abrir";
            btn.Name = "bt_AbrirOcorrencia";
            btn.UseColumnTextForButtonValue = true;

            /*if (!dr.HasRows)
                semOcorrencia();
            else
                panel_MO_NotFound.Visible = false;*/

            dr.Close();

            //FECHA CONEXÃO
            bdConn.Close();
        }

        //SEM OCORRÊNCIA
        void semOcorrencia()
        {
            if (this.dataGrid_MinhasOcorrencias.DataSource != null)
                this.dataGrid_MinhasOcorrencias.DataSource = null;
            else
            {
                this.dataGrid_MinhasOcorrencias.Rows.Clear();
                this.dataGrid_MinhasOcorrencias.Columns.Clear();
            }

            panel_MO_NotFound.Visible = true;
        }

        //CLICK DATAGRIDVIEW - BOTÃO ABRIR
        private void dataGrid_MinhasOcorrencias_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGrid_MinhasOcorrencias.Columns[11].Index)
            {
                if (dataGrid_MinhasOcorrencias.CurrentRow != null)
                    abreDetalhamentoOcorrencia(dataGrid_MinhasOcorrencias.CurrentRow.Cells[3].Value.ToString(), "Minhas Ocorrencia");
            }
        }

        //BOTÃO RETURN - MINHAS OCORRÊNCIAS
        private void bt_MinhasOcorrencias_Return_Click(object sender, EventArgs e)
        {
            inicioMenuOcorrencias();
        }

        //FORMATA CELULAS
        private void dataGrid_MinhasOcorrencias_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //STATUS GERAL
            /*if (e.Value != null && e.ColumnIndex == 6)
            {
                if (e.Value.Equals("Finalizado"))
                    e.CellStyle.BackColor = Color.DarkSeaGreen;
                else if (e.Value.Equals("Pendente"))
                    e.CellStyle.BackColor = Color.Khaki;
                else
                    e.CellStyle.BackColor = Color.Gray;
            }

            //STATUS RESPOSTA
            if (e.Value != null && e.ColumnIndex == 8)
            {
                if (e.Value.Equals("Aberta"))
                    e.CellStyle.BackColor = Color.Khaki;

                if (e.Value.Equals("Rejeitada") || e.Value.Equals("Cancelada"))
                    e.CellStyle.BackColor = Color.IndianRed;

                if (e.Value.Equals("Respondida") || e.Value.Equals("Resolvida"))
                    e.CellStyle.BackColor = Color.DarkSeaGreen;

                if (e.Value.Equals("Reaberta"))
                    e.CellStyle.BackColor = Color.Yellow;
            }*/

            //DATA PREVISTA SOLUÇÃO
            if (e.Value != null && e.ColumnIndex == 6)
                if (e.Value.Equals("Alta"))
                    e.CellStyle.BackColor = Color.IndianRed;
                else if (e.Value.Equals("Média"))
                    e.CellStyle.BackColor = Color.Yellow;
                else if (e.Value.Equals("Baixa"))
                    e.CellStyle.BackColor = Color.LemonChiffon;

            //DATA PREVISTA SOLUÇÃO
            if (e.Value != null && e.ColumnIndex == 9)
                if (e.Value.Equals(""))
                    e.Value = "-";
        }

        #endregion

        #region //****************************************** BUSCAR POR PROJETO  ******************************************\\

        //BOTÃO RETURN - BUSCAR POR PROJETO
        private void bt_OP_Return_Click(object sender, EventArgs e)
        {
            inicioMenuOcorrencias();
        }

        //TEXTBOX PESQUISA PROJETOS - BUSCAR POR PROJETO
        private void tb_OP_SelectPRJ_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (tb_OP_SelectPRJ.Text != "")
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    bdDataSet = new DataSet();
                    bdAdapter = new MySqlDataAdapter("SELECT cod_prj FROM projeto WHERE status_prj = 'Construção' AND cod_prj like '%" + tb_OP_SelectPRJ.Text + "%';", bdConn);
                    bdAdapter.Fill(bdDataSet, "projeto");
                    dataGrid_OP_SelectPRJ.DataSource = bdDataSet;

                    if (bdDataSet.Tables["projeto"].Rows.Count == 0)
                        lb_OP_NotFound.Visible = true;
                    else
                        lb_OP_NotFound.Visible = false;

                    dataGrid_OP_SelectPRJ.DataMember = "projeto";

                    //FECHA CONEXÃO
                    bdConn.Close();
                }
                else
                {
                    lb_OP_NotFound.Visible = false;
                    if (this.dataGrid_OP_SelectPRJ.DataSource != null)
                        this.dataGrid_OP_SelectPRJ.DataSource = null;
                    else
                    {
                        this.dataGrid_OP_SelectPRJ.Rows.Clear();
                        this.dataGrid_OP_SelectPRJ.Columns.Clear();
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //BOTÃO OK - SELECIONAR PROJETO - BUSCAR POR PROJETO
        private void bt_OP_OK_Click(object sender, EventArgs e)
        {
            if (dataGrid_OP_SelectPRJ.CurrentRow != null)
                if (bdDataSet.Tables["projeto"].Rows.Count > 0)
                {
                    try
                    {
                        abrePesquisaPRJ(dataGrid_OP_SelectPRJ.CurrentRow.Cells[0].Value.ToString());

                        lb_OP_PRJFiltros.Text = dataGrid_OP_SelectPRJ.CurrentRow.Cells[0].Value.ToString();

                        panel_MO_BuscarProjetos2.BringToFront();
                        panel_MO_BuscarProjetos2.Dock = DockStyle.Fill;
                        panel_MO_BuscarProjetos2.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        bdConn.Close();
                    }
                }
                else
                    MessageBox.Show("Nenhum projeto foi encontrado! Pesquise novamente.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //ARIR BUSCA POR PROJETO
        void abrePesquisaPRJ(string prj)
        {
            if (this.dataGrid_OP.DataSource != null)
                this.dataGrid_OP.DataSource = null;
            else
            {
                this.dataGrid_OP.Rows.Clear();
                this.dataGrid_OP.Columns.Clear();
            }

            //HEADER DATAGRIDVIEW
            dataGrid_OP.Columns.Add("cod_prj", "Projeto");
            dataGrid_OP.Columns.Add("dfrqf", "DF/RQF");
            dataGrid_OP.Columns.Add("sistema", "Sistema");
            dataGrid_OP.Columns.Add("cod_oco", "Cód. Ocorrência");
            dataGrid_OP.Columns.Add("identificador", "Identificador");
            dataGrid_OP.Columns.Add("classificacao", "Classificação");
            dataGrid_OP.Columns.Add("impacto_cttu", "Impacto CTTU");
            dataGrid_OP.Columns.Add("status_geral", "Status Geral");
            dataGrid_OP.Columns.Add("atrib_para", "Atribuído Para");
            dataGrid_OP.Columns.Add("status_resp", "Status da Resposta");
            dataGrid_OP.Columns.Add("dt_prev_sol", "DT Prevista Solução");

            bdConn.Open();
            MySqlCommand command = new MySqlCommand(criaQueryOP(prj), bdConn);
            MySqlDataReader dr = command.ExecuteReader();

            while (dr.Read())
                dataGrid_OP.Rows.Add(
                    dr["cod_prj"].ToString(),
                    dr["cod_rqf"].ToString(),
                    dr["sistema"].ToString(),
                    dr["cod_oco"].ToString(),
                    dr["identificador"].ToString(),
                    dr["classificacao"].ToString(),
                    dr["impacto_cttu"].ToString(),
                    dr["status_geral"].ToString(),
                    dr["atrib_para"].ToString(),
                    dr["status_resposta"].ToString(),
                    ((dr["dt_prev_solucao"].ToString() != "") ? String.Format("{0:dd/MM/yyyy}", DateTime.Parse(dr["dt_prev_solucao"].ToString())) : "-")
                    );

            //BOTÃO ABRIR
            DataGridViewButtonColumn btnOP = new DataGridViewButtonColumn();
            dataGrid_OP.Columns.Add(btnOP);
            btnOP.HeaderText = "Alteração";
            btnOP.Text = "Abrir";
            btnOP.Name = "bt_AbrirOcorrencia";
            btnOP.UseColumnTextForButtonValue = true;

            if (dr.HasRows)
                panel_OP_NotFound.Visible = false;
            else
            {
                if (this.dataGrid_OP.DataSource != null)
                    this.dataGrid_OP.DataSource = null;
                else
                {
                    this.dataGrid_OP.Rows.Clear();
                    this.dataGrid_OP.Columns.Clear();
                }

                panel_OP_NotFound.Visible = true;
            }

            //FECHA CONEXÃO
            bdConn.Close();
        }

        //CRIA QUERY BUSCA POR PRJETO
        string criaQueryOP(string projeto)
        {
            string query = "SELECT * FROM ocorrencia WHERE cod_prj = '" + projeto + "'" +
                ((tb_OP_RQF.Text != "") ? " AND cod_rqf like '%" + tb_OP_RQF.Text + "%'" : " AND cod_rqf like '%R%'") +
                ((cb_OP_Sistema.Text != "") ? " AND sistema like '%" + cb_OP_Sistema.Text + "%'" : "") +
                ((tb_OP_Identificador.Text != "") ? " AND identificador like '%" + tb_OP_Identificador.Text + "%'" : "") +
                ((tb_OP_RespFuncional.Text != "") ? " AND atrib_para like '%" + tb_OP_RespFuncional.Text + "%'" : "") +
                ((cb_OP_StatusGeral.Text != "") ? " AND status_geral = '" + cb_OP_StatusGeral.Text + "'" : "") +
                ((cb_OP_StatusResposta.Text != "") ? " AND status_resposta = '" + cb_OP_StatusResposta.Text + "'" : "")
                ;

            return query;
        }

        //CLICK DATAGRIDVIEW - BOTÃO ABRIR
        private void dataGrid_OP_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dataGrid_OP.Columns[11].Index)
            {
                if (dataGrid_OP.CurrentRow != null)
                    abreDetalhamentoOcorrencia(dataGrid_OP.CurrentRow.Cells[3].Value.ToString(), "Por Projeto");
            }
        }

        //BOTÃO RETURN - BUSCAR POR PROJETO
        private void bt_OP2_Return_Click(object sender, EventArgs e)
        {
            panel_MO_BuscarProjetos1.BringToFront();
            panel_MO_BuscarProjetos1.Visible = true;
            panel_MO_BuscarProjetos2.Visible = false;

            tb_OP_SelectPRJ.Text = "";
        }

        //FORMATA CELULAS
        private void dataGrid_OP_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            //STATUS GERAL
            /*if (e.Value != null && e.ColumnIndex == 6)
            {
                if (e.Value.Equals("Finalizado"))
                    e.CellStyle.BackColor = Color.DarkSeaGreen;
                else if (e.Value.Equals("Pendente"))
                    e.CellStyle.BackColor = Color.Khaki;
                else
                    e.CellStyle.BackColor = Color.Gray;
            }

            //STATUS RESPOSTA
            if (e.Value != null && e.ColumnIndex == 8)
            {
                if (e.Value.Equals("Aberta"))
                    e.CellStyle.BackColor = Color.Khaki;

                if (e.Value.Equals("Rejeitada") || e.Value.Equals("Cancelada"))
                    e.CellStyle.BackColor = Color.IndianRed;

                if (e.Value.Equals("Respondida") || e.Value.Equals("Resolvida"))
                    e.CellStyle.BackColor = Color.DarkSeaGreen;

                if (e.Value.Equals("Reaberta"))
                    e.CellStyle.BackColor = Color.Yellow;
            }*/

            //DATA PREVISTA SOLUÇÃO
            if (e.Value != null && e.ColumnIndex == 6)
                if (e.Value.Equals("Alta"))
                    e.CellStyle.BackColor = Color.IndianRed;
                else if (e.Value.Equals("Média"))
                    e.CellStyle.BackColor = Color.Yellow;
                else if (e.Value.Equals("Baixa"))
                    e.CellStyle.BackColor = Color.LemonChiffon;

            //DATA PREVISTA SOLUÇÃO
            if (e.Value != null && e.ColumnIndex == 9)
                if (e.Value.Equals(""))
                    e.Value = "-";
        }

        //DUPLO CLICK NA CELULA CELECIONADA - BUSCAR POR PROJETO
        private void dataGrid_OP_SelectPRJ_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGrid_OP_SelectPRJ.CurrentRow != null)
                bt_OP_OK.PerformClick();
            else
                MessageBox.Show("Selecionar um projeto da lista antes de prosseguir.", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //BOTÃO FILTROS - BUSCAR POR PROJETO
        private void bt_OP_Filtros_Click(object sender, EventArgs e)
        {
            if (!gb_OP_Filtros.Visible)
            {
                gb_OP_Filtros.Visible = true;

                bdConn.Open();

                MySqlCommand command = new MySqlCommand("SELECT nome_status FROM status_geral;", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                cb_OP_StatusGeral.Items.Clear();
                cb_OP_StatusGeral.Items.Add("");
                while (dr.Read())
                    cb_OP_StatusGeral.Items.Add(dr["nome_status"].ToString());
                dr.Close();

                command = new MySqlCommand("SELECT nome_status FROM status_resposta;", bdConn);
                dr = command.ExecuteReader();
                cb_OP_StatusResposta.Items.Clear();
                cb_OP_StatusResposta.Items.Add("");
                while (dr.Read())
                    cb_OP_StatusResposta.Items.Add(dr["nome_status"].ToString());
                dr.Close();

                bdConn.Close();
            }
            else
            {
                gb_OP_Filtros.Visible = false;

                tb_OP_RQF.Text = "";
                cb_OP_Sistema.Text = null;
                tb_OP_Identificador.Text = "";
                tb_OP_RespFuncional.Text = "";
                cb_OP_StatusGeral.Text = null;
                cb_OP_StatusResposta.Text = null;

                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
            }
        }

        #region FILTROS

        //RQF
        private void tb_OP_RQF_TextChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        //SISTEMA
        private void cb_OP_Sistema_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        //IDENTIFICADOR
        private void tb_OP_Identificador_TextChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        //RESPONSAVEL FUNCIONAL
        private void tb_OP_RespFuncional_TextChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        //STATUS GERAL
        private void cb_OP_StatusGeral_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        //STATUS RESPOSTA
        private void cb_OP_StatusResposta_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (gb_OP_Filtros.Visible)
                abrePesquisaPRJ(lb_OP_PRJFiltros.Text);
        }

        #endregion

        #endregion

        #region //****************************************** BUSCAR POR ID ******************************************\\

        //BOTÃO OK - OCORRENCIA POR ID
        private void bt_OID_SelectID_Click(object sender, EventArgs e)
        {
            if (tb_OID_SelectID.Text != "P_____-OC___")
            {
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT cod_oco FROM ocorrencia WHERE cod_oco = '" + tb_OID_SelectID.Text + "'", bdConn);
                MySqlDataReader dr = command.ExecuteReader();

                if (dr.HasRows)
                {
                    dr.Close();
                    bdConn.Close();
                    abreDetalhamentoOcorrencia(tb_OID_SelectID.Text, "Por ID");
                }
                else
                {
                    dr.Close();
                    bdConn.Close();
                    MessageBox.Show("Não existe nenhuma ocorrência com o código: " + tb_OID_SelectID.Text, "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tb_OID_SelectID.Text = "P_____-OC___";
                    tb_OID_SelectID.ForeColor = Color.Gray;
                }
            }
            else
                MessageBox.Show("Digite o código da ocorrência para continuar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        //BOTÃO RETURN - BUSCAR POR ID
        private void tb_OID_Return_Click(object sender, EventArgs e)
        {
            inicioMenuOcorrencias();
        }

        //CONTROLE TB SELECIONA ID
        private void tb_OID_SelectID_Enter(object sender, EventArgs e)
        {
            tb_OID_SelectID.Text = "";
            tb_OID_SelectID.ForeColor = Color.Black;
        }

        //CONTROLE TB SELECIONA ID
        private void tb_OID_SelectID_Leave(object sender, EventArgs e)
        {
            if (tb_OID_SelectID.Text == "")
            {
                tb_OID_SelectID.Text = "P_____-OC___";
                tb_OID_SelectID.ForeColor = Color.Gray;
            }
        }

        //VERICA CODIGO
        bool verificaCodID()
        {
            try
            {
                if (tb_OID_SelectID.Text.Substring(0, 1) != "P")
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(1, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(1, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(2, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(2, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(3, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(3, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(4, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(4, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(5, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(5, 1)) > 9)
                    return false;

                if (tb_OID_SelectID.Text.Substring(6, 1) != "-")
                    return false;

                if (tb_OID_SelectID.Text.Substring(7, 1) != "O")
                    return false;

                if (tb_OID_SelectID.Text.Substring(8, 1) != "C")
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(9, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(9, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(10, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(10, 1)) > 9)
                    return false;

                if (Int16.Parse(tb_OID_SelectID.Text.Substring(11, 1)) < 0 && Int16.Parse(tb_OID_SelectID.Text.Substring(11, 1)) > 9)
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        //PRESS ENTER - OCORRENCIAS POR ID
        private void tb_OID_SelectID_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
                bt_OID_SelectID.PerformClick();
        }

        #endregion

        #endregion        
    }
}
