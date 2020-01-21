using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleOcorrencias
{
    public partial class form_AddDataPrev : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        DateTime data;
        private string codigoOcorrencia;
        private string User;

        public form_AddDataPrev(string codOco, string user)
        {
            InitializeComponent();           
            codigoOcorrencia = codOco;
            User = user;
            calendar_DataPrev.MinDate = DateTime.Today;
        }

        private void bt_Adicionar_Click(object sender, EventArgs e)
        {          
            try
            {
                data = DateTime.Parse(calendar_DataPrev.SelectionRange.Start.ToShortDateString().ToString());
                string dt = String.Format("{0:yyyy-MM-dd}", data);                               

                bdConn.Open();

                MySqlCommand command = new MySqlCommand("UPDATE ocorrencia SET dt_prev_solucao = '" + dt + "' WHERE cod_oco = '" + codigoOcorrencia + "';", bdConn);
                command.ExecuteNonQuery();

                command = new MySqlCommand("INSERT INTO desc_ocorrencia (cod_oco, analista, tipo_registro, acao_analista) VALUES ('" + codigoOcorrencia + "','" + User + "', 'Atualização Data Prevista Solução', 'alterou a Data Prevista de Solução para (" + calendar_DataPrev.SelectionRange.Start.ToShortDateString() + ")');", bdConn);
                command.ExecuteNonQuery();

                MessageBox.Show("A Data Prevista para Solução foi alterada para o dia (" + calendar_DataPrev.SelectionRange.Start.ToShortDateString() + ").", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show("Erro ao atualizar Data Prevista de Solução.\n\nDetalhes: " + ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                bdConn.Close();
            }            
        }        
    }
}
