using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleDuvidas
{
    public partial class form_StatusGeral : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private string ocorrencia;
        private string analista;

        public form_StatusGeral(string codOcorrencia, string nomeAnalista)
        {
            InitializeComponent();
            this.ocorrencia = codOcorrencia;
            this.analista = nomeAnalista;
        }

        private void form_StatusGeral_Load(object sender, EventArgs e)
        {
            try
            {
                Form1 f1 = new Form1();
                string equipe = ((f1.Equipe) == "Fábrica Desenvolvimento" ? "FD" : "FF");


                //ABRE CONEXÃO
                bdConn.Open();

                //EXECUTA COMANDO
                MySqlCommand command = new MySqlCommand("SELECT * FROM status_geral WHERE equipe like '%" + equipe + "%';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();

                while (dr.Read())
                    cb_SelecionarStatus.Items.Add(dr["nome_status"].ToString());

                //FECHA CONEXÃO
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        private void bt_Cancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bt_Atualizar_Click(object sender, EventArgs e)
        {
            if (cb_SelecionarStatus.Text != "")
            {
                try
                {
                    //ABRE CONEXÃO
                    bdConn.Open();

                    //EXECUTA COMANDO
                    MySqlCommand command = new MySqlCommand("UPDATE ocorrencia SET status_geral = '" + cb_SelecionarStatus.Text + "' WHERE cod_oco = '" + ocorrencia + "';", bdConn);
                    command.ExecuteNonQuery();

                    command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + ocorrencia + "', '" + analista + "', 'Atualizacao Status', 'atualizou o Status Geral para " + cb_SelecionarStatus.Text + "');", bdConn);
                    command.ExecuteNonQuery();

                    //FECHA CONEXÃO
                    bdConn.Close();

                    MessageBox.Show("Status atualizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    this.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bdConn.Close();
                }
            }
            else
                MessageBox.Show("Selecione um Status para atualizar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
    }
}
