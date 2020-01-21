using System;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleOcorrencias
{
    public partial class form_RelatorioGestao : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        DataSet bdDataSet;
        MySqlDataAdapter bdAdapter;
        private string usuario;
        private string opcao;

        public form_RelatorioGestao(string user, string opcao)
        {
            InitializeComponent();
            this.usuario = user;
            this.opcao = opcao;
        }

        private void form_RelatorioGestao_Load(object sender, EventArgs e)
        {
            switch (opcao)
            {
                case "Por Status":
                    panel_PRJ.SendToBack();
                    panel_PRJ.Visible = false;
                    label1.Text = "Quantidade de Ocorrências (Projeto x Status Geral):";
                    bt_Voltar.Visible = false;
                    inicioPorStatus();
                    break;
                case "Por Ocorrência":
                    panel_PRJ.BringToFront();
                    panel_PRJ.Dock = DockStyle.Fill;
                    panel_PRJ.Visible = true;
                    bt_Voltar.Visible = true;
                    break;
            }
        }

        void inicioPorStatus()
        {
            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                #region RECUPERA VIEW RELATORIO

                MySqlCommand comand = new MySqlCommand("SELECT view_relatorio FROM usuarios WHERE user_nome = '" + usuario + "';", bdConn);
                MySqlDataReader dr = comand.ExecuteReader();
                while (dr.Read())
                    cb_ViewRelatorio.Checked = (dr["view_relatorio"].ToString() == "sim") ? true : false;
                dr.Close();

                #endregion

                #region RECUPERA STATUS

                List<string> listStatus = new List<string>();
                comand = new MySqlCommand("SELECT nome_status FROM status_geral ORDER BY nome_status;", bdConn);
                dr = comand.ExecuteReader();
                while (dr.Read())
                    listStatus.Add(dr["nome_status"].ToString());
                dr.Close();

                comand = new MySqlCommand("SELECT nome_status FROM status_impacto ORDER BY nome_status;", bdConn);
                dr = comand.ExecuteReader();
                while (dr.Read())
                    listStatus.Add(dr["nome_status"].ToString());
                dr.Close();

                #endregion

                #region RECUPERA PROJETOS

                List<string> listProjetos = new List<string>();
                comand = new MySqlCommand("SELECT cod_prj FROM projeto WHERE status_prj = 'Construção' ORDER BY cod_prj;", bdConn);
                dr = comand.ExecuteReader();
                while (dr.Read())
                    listProjetos.Add(dr[0].ToString());
                dr.Close();

                #endregion

                //LIMPA GRIDVIEW
                if (this.dataGrid_TotalizadorPorProjeto.DataSource != null)
                    this.dataGrid_TotalizadorPorProjeto.DataSource = null;
                else
                {
                    this.dataGrid_TotalizadorPorProjeto.Rows.Clear();
                    this.dataGrid_TotalizadorPorProjeto.Columns.Clear();
                }

                //HEADER DATAGRIDVIEW
                dataGrid_TotalizadorGeral.Columns.Add("cod_prj", "PROJETO");
                dataGrid_TotalizadorPorProjeto.Columns.Add("cod_prj", "PROJETO");
                foreach (var status in listStatus)
                {
                    dataGrid_TotalizadorGeral.Columns.Add(status, status.ToUpper());
                    dataGrid_TotalizadorPorProjeto.Columns.Add(status, status.ToUpper());
                }


                int coluna = 0;

                dataGrid_TotalizadorGeral.Rows.Add();
                dataGrid_TotalizadorGeral.Rows[0].Cells[0].Value = "TOTAL:";

                foreach (var status in listStatus)
                {
                    coluna++;
                    comand = new MySqlCommand("SELECT count(cod_oco) FROM ocorrencia WHERE status_geral = '" + status + "';", bdConn);
                    dr = comand.ExecuteReader();
                    if (dr.Read())
                        dataGrid_TotalizadorGeral.Rows[0].Cells[coluna].Value = dr[0].ToString();
                    dr.Close();
                }


                int linha = 0;
                foreach (var projeto in listProjetos)
                {
                    coluna = 0;
                    dataGrid_TotalizadorPorProjeto.Rows.Add();
                    dataGrid_TotalizadorPorProjeto.Rows[linha].Cells[coluna].Value = projeto;

                    foreach (var status in listStatus)
                    {
                        coluna++;
                        comand = new MySqlCommand("SELECT count(cod_oco) FROM ocorrencia WHERE cod_prj = '" + projeto + "' AND status_geral = '" + status + "';", bdConn);
                        dr = comand.ExecuteReader();
                        if (dr.Read())
                            dataGrid_TotalizadorPorProjeto.Rows[linha].Cells[coluna].Value = dr[0].ToString();
                        dr.Close();
                    }

                    linha++;
                }

                dataGrid_TotalizadorPorProjeto.ClearSelection();
                dataGrid_TotalizadorGeral.ClearSelection();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //FECHA CONEXÃO
                bdConn.Close();
            }
        }

        void inicioPorOcorrencia(string projeto)
        {
            label1.Text = "Relatótio de Ocorrências (" + projeto + "):";

            try
            {
                //ABRE CONEXÃO
                bdConn.Open();

                List<string> listOcocorrencia = new List<string>();
                string status_geral = "";
                DateTime data_abertura = DateTime.Now;
                DateTime ult_alteração = DateTime.Now;
                int row = 0;

                #region RECUPERA VIEW RELATORIO

                MySqlCommand comand = new MySqlCommand("SELECT view_relatorio FROM usuarios WHERE user_nome = '" + usuario + "';", bdConn);
                MySqlDataReader dr = comand.ExecuteReader();
                if (dr.Read())
                    cb_ViewRelatorio.Checked = (dr["view_relatorio"].ToString() == "sim") ? true : false;
                dr.Close();

                #endregion                
              
                //LIMPA GRIDVIEW
                if (this.dataGrid_TotalizadorPorProjeto.DataSource != null)
                    this.dataGrid_TotalizadorPorProjeto.DataSource = null;
                else
                {
                    this.dataGrid_TotalizadorPorProjeto.Rows.Clear();
                    this.dataGrid_TotalizadorPorProjeto.Columns.Clear();
                }

                //HEADER DATAGRIDVIEW
                dataGrid_TotalizadorPorProjeto.Columns.Add("cod_prj", "PROJETO");
                dataGrid_TotalizadorPorProjeto.Columns.Add("cod_oco", "OCORRÊNCIA");
                dataGrid_TotalizadorPorProjeto.Columns.Add("status_geral", "STATUS GERAL");
                dataGrid_TotalizadorPorProjeto.Columns.Add("data_abertura", "DATA ABERTURA");
                dataGrid_TotalizadorPorProjeto.Columns.Add("ult_alteracao", "ÚLTIMA ALTERAÇÃO");
                dataGrid_TotalizadorPorProjeto.Columns.Add("tempo_paralisado", "TEMPO PARALISADO");

                #region CARREGA OCORRENCIAS DO PROJETO

                comand = new MySqlCommand("SELECT cod_oco FROM ocorrencia WHERE cod_prj = '" + projeto + "' AND status_geral != 'Finalizado' AND status_geral != 'Cancelado';", bdConn);
                dr = comand.ExecuteReader();
                while (dr.Read())
                    listOcocorrencia.Add(dr["cod_oco"].ToString());                
                dr.Close();

                #endregion                         

                foreach (var oco in listOcocorrencia)
                {
                    comand = new MySqlCommand("SELECT A.status_geral, IF ( A.horario_registro < B.ultima_data, B.ultima_data, A.horario_registro ) AS DATA_RESULT FROM(SELECT cod_oco, cod_prj, status_geral, horario_registro FROM ocorrencia) AS A JOIN(SELECT cod_oco, max(horario_registro) AS ultima_data FROM desc_ocorrencia GROUP BY cod_oco) AS B ON A.cod_oco = '" + oco + "' AND B.cod_oco = '" + oco + "';", bdConn);
                    dr = comand.ExecuteReader();
                    if (dr.Read())
                    {
                        status_geral = dr["status_geral"].ToString();
                        ult_alteração = DateTime.Parse(dr["DATA_RESULT"].ToString());
                    }                        
                    dr.Close();

                    comand = new MySqlCommand("SELECT horario_registro FROM  desc_ocorrencia WHERE tipo_registro = 'Abertura Ocorrencia' AND cod_oco = '" + oco + "';", bdConn);
                    dr = comand.ExecuteReader();
                    if (dr.Read())                    
                        data_abertura = DateTime.Parse(dr["horario_registro"].ToString());                    
                    dr.Close();

                    TimeSpan tempo_parado = DateTime.Now - ult_alteração;

                    dataGrid_TotalizadorPorProjeto.Rows.Add();                    
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[0].Value = projeto;
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[1].Value = oco;
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[2].Value = status_geral;
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[3].Value = String.Format("{0:dd/MM/yyyy}", data_abertura);
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[4].Value = String.Format("{0:dd/MM/yyyy}", ult_alteração);
                    dataGrid_TotalizadorPorProjeto.Rows[row].Cells[5].Value = ((tempo_parado.Days > 1) ? (tempo_parado.Days.ToString() + " Dias") : (tempo_parado.Days.ToString() + " Dia"));
                    row++;
                }
                
                dataGrid_TotalizadorPorProjeto.ClearSelection();               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                //FECHA CONEXÃO
                bdConn.Close();
            }
        }

        private void bt_Ok_Click(object sender, EventArgs e)
        {
            try
            {
                bdConn.Open();

                MySqlCommand command = new MySqlCommand("UPDATE usuarios SET view_relatorio = '" + ((cb_ViewRelatorio.Checked == true) ? "sim" : "nao") + "' WHERE user_nome = '" + usuario + "';", bdConn);
                command.ExecuteNonQuery();
            }
            catch
            {
            }
            finally
            {
                bdConn.Close();
            }
            
            this.Close();
        }

        private void dataGrid_TotalizadorGeral_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if ((e.Value != null) && (e.ColumnIndex == 0))
            {
                e.CellStyle.Font = new Font(e.CellStyle.Font, FontStyle.Bold);
            }       
        }

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

        private void bt_OP_OK_Click(object sender, EventArgs e)
        {
            if (dataGrid_OP_SelectPRJ.CurrentRow != null)
                if (bdDataSet.Tables["projeto"].Rows.Count > 0)
                {
                    try
                    {
                        panel_PRJ.Visible = false;
                        inicioPorOcorrencia(dataGrid_OP_SelectPRJ.CurrentRow.Cells[0].Value.ToString());
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

        private void dataGrid_OP_SelectPRJ_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            bt_OP_OK.PerformClick();
        }

        private void bt_Voltar_Click(object sender, EventArgs e)
        {
            panel_PRJ.BringToFront();
            panel_PRJ.Visible = true;
            tb_OP_SelectPRJ.Text = "";
        }
    }
}
