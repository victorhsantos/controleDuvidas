using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleOcorrencias
{
    public partial class form_AddAcompanharPRJ : Form
    {
        //BANCO DE DADOS
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private static string NomeSelecionados;

        public string nomeSelecionados
        {
            get {return NomeSelecionados; }
        }

        public form_AddAcompanharPRJ()
        {
            InitializeComponent();
            NomeSelecionados = "";

            try
            {
                clb_Usuarios.Items.Clear();
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT user_nome FROM usuarios ORDER BY user_nome;", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                while (dr.Read())
                    clb_Usuarios.Items.Add(dr["user_nome"].ToString());
                dr.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro inesperado ao listar usuários.\n\nDetalhes: " + ex.ToString(), "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();
            }
            finally
            {
                bdConn.Close();
            }
        }
         
        private void bt_Cancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bt_Adicionar_Click(object sender, EventArgs e)
        {
            if (clb_Usuarios.CheckedItems.Count != 0)
            {
                foreach (object item in clb_Usuarios.CheckedItems)
                    NomeSelecionados += ((NomeSelecionados == "") ? item : (", " + item));
                this.Close();
            }
            else
            {
                MessageBox.Show("Selecione os usuários antes de adicionar!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
