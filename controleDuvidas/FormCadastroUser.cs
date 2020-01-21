using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace controleDuvidas
{
    public partial class FormCadastroUser : Form
    {
        //BANCO DE DADOS
        private MySqlConnection bdConn = new MySqlConnection(" Persist Security Info=False;server=192.168.10.6;database=controle_ocorrencias;uid=admin;server = 192.168.10.6; database = controle_ocorrencias; uid = admin; pwd = accenture; Allow Zero Datetime=True");

        public FormCadastroUser()
        {
            InitializeComponent();
        }

        //LOAD FROM
        private void FormCadastroUser_Load(object sender, EventArgs e)
        {
            tb_Cadastro_User.Text = Environment.UserName;
            tb_Cadastro_Email.Text = Environment.UserName + "@accenture.com";
            tb_Cadastro_User.ForeColor = Color.Black;
            tb_Cadastro_Email.ForeColor = Color.Black;
        }

        //BOTÃO OK
        private void bt_CadastroUser_Ok_Click(object sender, EventArgs e)
        {
            try
            {
                if (verificaCampos())
                    if (!lb_NomeUserExistente.Visible)
                        if (!lb_UserExistente.Visible)
                            if (tb_Cadastro_Password.Text.Length > 3)
                                if (verificaEmail(tb_Cadastro_Email.Text))
                                {
                                    //ABRE CONEXÃO
                                    bdConn.Open();

                                    MySqlCommand command = new MySqlCommand("INSERT INTO usuarios (user_login, user_senha, user_nome, user_email, user_equipe, user_lvl, permissao_relatorio, view_relatorio) VALUES ('" + tb_Cadastro_User.Text + "', '" + tb_Cadastro_Password.Text + "', '" + tb_Cadastro_Nome.Text + "', '" + tb_Cadastro_Email.Text + "', '" + cb_Cadastro_Equipe.Text + "', " + "'0', 'nao', 'nao'" + ");", bdConn);
                                    command.ExecuteNonQuery();

                                    MessageBox.Show("O cadastro foi realizado com sucesso!", "Concluído!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                    resertaCampos();

                                    //FECHA CONEXÃO
                                    bdConn.Close();

                                    this.Close();
                                }
                                else
                                    MessageBox.Show("Email Incorreto!", "Preenchimento Obrigatório!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            else
                                MessageBox.Show("A senha deve ter no minimo 4 caracteres!", "Atenção!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        else
                            MessageBox.Show("Usuário já existe!", "Preenchimento Obrigatório!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    else
                        MessageBox.Show("Nome de Usuário já existe!", "Preenchimento Obrigatório!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                else
                    MessageBox.Show("Favor preencher todos os campos!", "Preenchimento Obrigatório!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            catch
            {
                MessageBox.Show("Este email já esta cadastrado!", "Erro!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                bdConn.Close();
            }
        }

        //VERIFICA CAMPOS PREENCHIDOS
        bool verificaCampos()
        {
            if (tb_Cadastro_Nome.Text == "Nome e Sobrenome")
                return false;

            if (tb_Cadastro_User.Text == "Usuário")
                return false;

            if (tb_Cadastro_Password.Text == "SenH@")
                return false;

            if (tb_Cadastro_Email.Text == "Email")
                return false;

            if (cb_Cadastro_Equipe.Text == "")
                return false;

            return true;
        }

        //VERIFICA EMAIL
        bool verificaEmail(string email)
        {            
            int indexArr = email.IndexOf('@');
            if (indexArr > 0)
            {
                int indexDot = email.IndexOf('.', indexArr);
                if (indexDot - 1 > indexArr)
                {
                    if (indexDot + 1 < email.Length)
                    {
                        string indexDot2 = email.Substring(indexDot + 1, 1);
                        if (indexDot2 != ".")
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        //RESETA CAMPOS
        void resertaCampos()
        {
            tb_Cadastro_Nome.ForeColor = Color.DarkGray;
            tb_Cadastro_User.ForeColor = Color.DarkGray;
            tb_Cadastro_Password.ForeColor = Color.DarkGray;
            tb_Cadastro_Email.ForeColor = Color.DarkGray;

            tb_Cadastro_Nome.Text = "Nome e Sobrenome";
            tb_Cadastro_User.Text = "Usuário";
            tb_Cadastro_Password.Text = "SenH@";
            tb_Cadastro_Email.Text = "Email";
            cb_Cadastro_Equipe.Text = null;

        }

        //VERIFICA SE USUÁRIO JÁ EXISTE
        private void tb_Cadastro_User_TextChanged(object sender, EventArgs e)
        {
            if (bdConn.State == ConnectionState.Closed)
            {
                bdConn.Open();

                MySqlCommand comand = new MySqlCommand("SELECT user_login FROM usuarios WHERE user_login = '" + tb_Cadastro_User.Text + "';", bdConn);
                MySqlDataReader dr = comand.ExecuteReader();

                if (dr.Read())
                    lb_UserExistente.Visible = true;
                else
                    lb_UserExistente.Visible = false;

                bdConn.Close();
            }            
        }

        //VERIFICA SE NOME DE USUÁRIO JÁ EXISTE
        private void tb_Cadastro_Nome_TextChanged(object sender, EventArgs e)
        {
            if (bdConn.State == ConnectionState.Closed)
            {
                bdConn.Open();

                MySqlCommand comand = new MySqlCommand("SELECT user_nome FROM usuarios WHERE user_nome = '" + tb_Cadastro_Nome.Text + "';", bdConn);
                MySqlDataReader dr = comand.ExecuteReader();

                if (dr.Read())
                    lb_NomeUserExistente.Visible = true;
                else
                    lb_NomeUserExistente.Visible = false;

                bdConn.Close();
            }            
        }

        //VERIFICA SEGURANÇA DA SENHA
        private void tb_Cadastro_Password_TextChanged(object sender, EventArgs e)
        {
            if (tb_Cadastro_Password.Text == "" || tb_Cadastro_Password.Text == "SenH@")
            {
                lb_Password.Visible = false;
            }
            else
            {
                lb_Password.Visible = true;

                if (tb_Cadastro_Password.Text.Length < 3)
                {
                    lb_Password.Text = "Segurança Fraca";
                    lb_Password.ForeColor = Color.Red;
                }
                else if (tb_Cadastro_Password.Text.Length > 3 && tb_Cadastro_Password.Text.Length < 8)
                {
                    lb_Password.Text = "Segurança Média";
                    lb_Password.ForeColor = Color.DarkOrange;
                }
                else if (tb_Cadastro_Password.Text.Length > 8)
                {
                    lb_Password.Text = "Segurança Alta";
                    lb_Password.ForeColor = Color.DarkGreen;
                }
            }
        }

        //BOTÃO SAIR
        private void bt_CadastroUser_Sair_Click(object sender, EventArgs e)
        {
            resertaCampos();
            this.Close();
        }        

        #region //****************************************** CAMPOS DEFAULT  ******************************************\\

        private void tb_Cadastro_Nome_Enter(object sender, EventArgs e)
        {
            if (tb_Cadastro_Nome.Text == "Nome e Sobrenome")
            {
                tb_Cadastro_Nome.Text = "";
                tb_Cadastro_Nome.ForeColor = Color.Black;
            }
        }

        private void tb_Cadastro_Nome_Leave(object sender, EventArgs e)
        {
            if (tb_Cadastro_Nome.Text == "")
            {
                tb_Cadastro_Nome.Text = "Nome e Sobrenome";
                tb_Cadastro_Nome.ForeColor = Color.DarkGray;
            }
        }

        private void tb_Cadastro_User_Enter(object sender, EventArgs e)
        {
            if (tb_Cadastro_User.Text == "Usuário")
            {
                tb_Cadastro_User.Text = "";
                tb_Cadastro_User.ForeColor = Color.Black;
            }
        }

        private void tb_Cadastro_User_Leave(object sender, EventArgs e)
        {
            if (tb_Cadastro_User.Text == "")
            {
                tb_Cadastro_User.Text = "Usuário";
                tb_Cadastro_User.ForeColor = Color.DarkGray;
            }
        }

        private void tb_Cadastro_Password_Enter(object sender, EventArgs e)
        {
            if (tb_Cadastro_Password.Text == "Senha")
            {
                tb_Cadastro_Password.Text = "";
                tb_Cadastro_Password.ForeColor = Color.Black;
            }
        }

        private void tb_Cadastro_Password_Leave(object sender, EventArgs e)
        {
            if (tb_Cadastro_Password.Text == "")
            {
                tb_Cadastro_Password.Text = "Senha";
                tb_Cadastro_Password.ForeColor = Color.DarkGray;
            }
        }

        private void tb_Cadastro_Email_Enter(object sender, EventArgs e)
        {
            if (tb_Cadastro_Email.Text == "Email")
            {
                tb_Cadastro_Email.Text = "";
                tb_Cadastro_Email.ForeColor = Color.Black;
            }
        }

        private void tb_Cadastro_Email_Leave(object sender, EventArgs e)
        {
            if (tb_Cadastro_Email.Text == "")
            {
                tb_Cadastro_Email.Text = "Email";
                tb_Cadastro_Email.ForeColor = Color.DarkGray;
            }
        }

        #endregion                      
      
    }
}
