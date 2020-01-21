using System;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleDuvidas
{
    public partial class form_AlterarSenha : Form
    {
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");
        private string user;

        public form_AlterarSenha(string usuario)
        {
            InitializeComponent();

            this.user = usuario;
            tb_Usuario.Text = usuario;
        }

        private void bt_Cancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void bt_Salvar_Click(object sender, EventArgs e)
        {
            if (verificaCampos())
                if (verificaSenhaAtual())
                {
                    bdConn.Open();
                    MySqlCommand command = new MySqlCommand("UPDATE usuarios SET user_senha = '" + tb_NovaSenha.Text + "' WHERE user_nome = '" + user + "';", bdConn);
                    command.ExecuteNonQuery();

                    bdConn.Close();

                    MessageBox.Show("Nova Senha cadastrada com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    limpaCampos();

                    this.Close();
                }
        }

        bool verificaCampos()
        {
            if (tb_SenhaAtual.Text == "")
            {
                MessageBox.Show("Inserir senha atual!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (tb_NovaSenha.Text == "")
            {
                MessageBox.Show("Colocar nova senha!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            if (tb_NovaSenha.Text.Length < 4)
            {
                MessageBox.Show("A senha deve ter no minimo 4 caracteres!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        bool verificaSenhaAtual()
        {
            try
            {
                bdConn.Open();
                MySqlCommand command = new MySqlCommand("SELECT user_senha FROM usuarios WHERE user_nome = '" + user + "';", bdConn);
                MySqlDataReader dr = command.ExecuteReader();
                if (dr.Read())
                    if (tb_SenhaAtual.Text != dr["user_senha"].ToString())
                    {
                        MessageBox.Show("Senha atual inválida!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        dr.Close();
                        bdConn.Close();
                        return false;
                    }
                dr.Close();
                bdConn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
            return true;
        }

        void limpaCampos()
        {
            tb_NovaSenha.Text = "";
            tb_SenhaAtual.Text = "";
        }

        private void tb_NovaSenha_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == 13)
                bt_Salvar.PerformClick();
        }
    }
}
