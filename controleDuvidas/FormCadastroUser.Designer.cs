namespace controleDuvidas
{
    partial class FormCadastroUser
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormCadastroUser));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lb_UserExistente = new System.Windows.Forms.Label();
            this.cb_Cadastro_Equipe = new System.Windows.Forms.ComboBox();
            this.tb_Cadastro_Nome = new System.Windows.Forms.TextBox();
            this.pictureBox5 = new System.Windows.Forms.PictureBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.tb_Cadastro_Email = new System.Windows.Forms.TextBox();
            this.tb_Cadastro_Password = new System.Windows.Forms.TextBox();
            this.tb_Cadastro_User = new System.Windows.Forms.TextBox();
            this.bt_CadastroUser_Ok = new System.Windows.Forms.Button();
            this.bt_CadastroUser_Sair = new System.Windows.Forms.Button();
            this.lb_NomeUserExistente = new System.Windows.Forms.Label();
            this.lb_Password = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(48, 70);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(51, 41);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
            this.pictureBox2.Location = new System.Drawing.Point(48, 117);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(51, 41);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox2.TabIndex = 1;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
            this.pictureBox3.Location = new System.Drawing.Point(48, 164);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(51, 41);
            this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox3.TabIndex = 2;
            this.pictureBox3.TabStop = false;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel1.Controls.Add(this.lb_Password);
            this.panel1.Controls.Add(this.lb_NomeUserExistente);
            this.panel1.Controls.Add(this.lb_UserExistente);
            this.panel1.Controls.Add(this.cb_Cadastro_Equipe);
            this.panel1.Controls.Add(this.tb_Cadastro_Nome);
            this.panel1.Controls.Add(this.pictureBox5);
            this.panel1.Controls.Add(this.pictureBox4);
            this.panel1.Controls.Add(this.tb_Cadastro_Email);
            this.panel1.Controls.Add(this.tb_Cadastro_Password);
            this.panel1.Controls.Add(this.tb_Cadastro_User);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Controls.Add(this.pictureBox3);
            this.panel1.Controls.Add(this.pictureBox2);
            this.panel1.Location = new System.Drawing.Point(1, 63);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(482, 275);
            this.panel1.TabIndex = 3;
            // 
            // lb_UserExistente
            // 
            this.lb_UserExistente.AutoSize = true;
            this.lb_UserExistente.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_UserExistente.ForeColor = System.Drawing.Color.Red;
            this.lb_UserExistente.Location = new System.Drawing.Point(105, 63);
            this.lb_UserExistente.Name = "lb_UserExistente";
            this.lb_UserExistente.Size = new System.Drawing.Size(89, 14);
            this.lb_UserExistente.TabIndex = 11;
            this.lb_UserExistente.Text = "Usuário já existe!";
            this.lb_UserExistente.Visible = false;
            // 
            // cb_Cadastro_Equipe
            // 
            this.cb_Cadastro_Equipe.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_Cadastro_Equipe.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cb_Cadastro_Equipe.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cb_Cadastro_Equipe.FormattingEnabled = true;
            this.cb_Cadastro_Equipe.Items.AddRange(new object[] {
            "Fábrica Desenvolvimento",
            "Fábrica Funcional"});
            this.cb_Cadastro_Equipe.Location = new System.Drawing.Point(105, 220);
            this.cb_Cadastro_Equipe.Name = "cb_Cadastro_Equipe";
            this.cb_Cadastro_Equipe.Size = new System.Drawing.Size(316, 26);
            this.cb_Cadastro_Equipe.TabIndex = 5;
            // 
            // tb_Cadastro_Nome
            // 
            this.tb_Cadastro_Nome.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Cadastro_Nome.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_Cadastro_Nome.ForeColor = System.Drawing.Color.DarkGray;
            this.tb_Cadastro_Nome.Location = new System.Drawing.Point(105, 32);
            this.tb_Cadastro_Nome.MaxLength = 54;
            this.tb_Cadastro_Nome.Name = "tb_Cadastro_Nome";
            this.tb_Cadastro_Nome.Size = new System.Drawing.Size(316, 25);
            this.tb_Cadastro_Nome.TabIndex = 1;
            this.tb_Cadastro_Nome.Text = "Nome e Sobrenome";
            this.tb_Cadastro_Nome.TextChanged += new System.EventHandler(this.tb_Cadastro_Nome_TextChanged);
            this.tb_Cadastro_Nome.Enter += new System.EventHandler(this.tb_Cadastro_Nome_Enter);
            this.tb_Cadastro_Nome.Leave += new System.EventHandler(this.tb_Cadastro_Nome_Leave);
            // 
            // pictureBox5
            // 
            this.pictureBox5.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox5.Image")));
            this.pictureBox5.Location = new System.Drawing.Point(48, 23);
            this.pictureBox5.Name = "pictureBox5";
            this.pictureBox5.Size = new System.Drawing.Size(51, 41);
            this.pictureBox5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox5.TabIndex = 8;
            this.pictureBox5.TabStop = false;
            // 
            // pictureBox4
            // 
            this.pictureBox4.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox4.Image")));
            this.pictureBox4.Location = new System.Drawing.Point(48, 211);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(51, 41);
            this.pictureBox4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox4.TabIndex = 6;
            this.pictureBox4.TabStop = false;
            // 
            // tb_Cadastro_Email
            // 
            this.tb_Cadastro_Email.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Cadastro_Email.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_Cadastro_Email.ForeColor = System.Drawing.Color.DarkGray;
            this.tb_Cadastro_Email.Location = new System.Drawing.Point(105, 171);
            this.tb_Cadastro_Email.MaxLength = 100;
            this.tb_Cadastro_Email.Name = "tb_Cadastro_Email";
            this.tb_Cadastro_Email.Size = new System.Drawing.Size(316, 25);
            this.tb_Cadastro_Email.TabIndex = 4;
            this.tb_Cadastro_Email.Text = "Email";
            this.tb_Cadastro_Email.Enter += new System.EventHandler(this.tb_Cadastro_Email_Enter);
            this.tb_Cadastro_Email.Leave += new System.EventHandler(this.tb_Cadastro_Email_Leave);
            // 
            // tb_Cadastro_Password
            // 
            this.tb_Cadastro_Password.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Cadastro_Password.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_Cadastro_Password.ForeColor = System.Drawing.Color.DarkGray;
            this.tb_Cadastro_Password.Location = new System.Drawing.Point(105, 126);
            this.tb_Cadastro_Password.MaxLength = 15;
            this.tb_Cadastro_Password.Name = "tb_Cadastro_Password";
            this.tb_Cadastro_Password.PasswordChar = '•';
            this.tb_Cadastro_Password.Size = new System.Drawing.Size(316, 25);
            this.tb_Cadastro_Password.TabIndex = 3;
            this.tb_Cadastro_Password.Text = "Senha";
            this.tb_Cadastro_Password.TextChanged += new System.EventHandler(this.tb_Cadastro_Password_TextChanged);
            this.tb_Cadastro_Password.Enter += new System.EventHandler(this.tb_Cadastro_Password_Enter);
            this.tb_Cadastro_Password.Leave += new System.EventHandler(this.tb_Cadastro_Password_Leave);
            // 
            // tb_Cadastro_User
            // 
            this.tb_Cadastro_User.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Cadastro_User.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tb_Cadastro_User.ForeColor = System.Drawing.Color.DarkGray;
            this.tb_Cadastro_User.Location = new System.Drawing.Point(105, 79);
            this.tb_Cadastro_User.MaxLength = 54;
            this.tb_Cadastro_User.Name = "tb_Cadastro_User";
            this.tb_Cadastro_User.Size = new System.Drawing.Size(316, 25);
            this.tb_Cadastro_User.TabIndex = 2;
            this.tb_Cadastro_User.Text = "Usuário";
            this.tb_Cadastro_User.TextChanged += new System.EventHandler(this.tb_Cadastro_User_TextChanged);
            this.tb_Cadastro_User.Enter += new System.EventHandler(this.tb_Cadastro_User_Enter);
            this.tb_Cadastro_User.Leave += new System.EventHandler(this.tb_Cadastro_User_Leave);
            // 
            // bt_CadastroUser_Ok
            // 
            this.bt_CadastroUser_Ok.BackColor = System.Drawing.Color.SeaGreen;
            this.bt_CadastroUser_Ok.FlatAppearance.BorderSize = 0;
            this.bt_CadastroUser_Ok.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_CadastroUser_Ok.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bt_CadastroUser_Ok.Location = new System.Drawing.Point(262, 344);
            this.bt_CadastroUser_Ok.Name = "bt_CadastroUser_Ok";
            this.bt_CadastroUser_Ok.Size = new System.Drawing.Size(135, 38);
            this.bt_CadastroUser_Ok.TabIndex = 0;
            this.bt_CadastroUser_Ok.Text = "Ok";
            this.bt_CadastroUser_Ok.UseVisualStyleBackColor = false;
            this.bt_CadastroUser_Ok.Click += new System.EventHandler(this.bt_CadastroUser_Ok_Click);
            // 
            // bt_CadastroUser_Sair
            // 
            this.bt_CadastroUser_Sair.BackColor = System.Drawing.Color.Firebrick;
            this.bt_CadastroUser_Sair.FlatAppearance.BorderSize = 0;
            this.bt_CadastroUser_Sair.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_CadastroUser_Sair.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bt_CadastroUser_Sair.Location = new System.Drawing.Point(119, 344);
            this.bt_CadastroUser_Sair.Name = "bt_CadastroUser_Sair";
            this.bt_CadastroUser_Sair.Size = new System.Drawing.Size(137, 38);
            this.bt_CadastroUser_Sair.TabIndex = 6;
            this.bt_CadastroUser_Sair.Text = "Sair";
            this.bt_CadastroUser_Sair.UseVisualStyleBackColor = false;
            this.bt_CadastroUser_Sair.Click += new System.EventHandler(this.bt_CadastroUser_Sair_Click);
            // 
            // lb_NomeUserExistente
            // 
            this.lb_NomeUserExistente.AutoSize = true;
            this.lb_NomeUserExistente.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_NomeUserExistente.ForeColor = System.Drawing.Color.Red;
            this.lb_NomeUserExistente.Location = new System.Drawing.Point(105, 15);
            this.lb_NomeUserExistente.Name = "lb_NomeUserExistente";
            this.lb_NomeUserExistente.Size = new System.Drawing.Size(134, 14);
            this.lb_NomeUserExistente.TabIndex = 12;
            this.lb_NomeUserExistente.Text = "Nome de Usuário já existe!";
            this.lb_NomeUserExistente.Visible = false;
            // 
            // lb_Password
            // 
            this.lb_Password.AutoSize = true;
            this.lb_Password.Font = new System.Drawing.Font("Arial", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lb_Password.ForeColor = System.Drawing.Color.Green;
            this.lb_Password.Location = new System.Drawing.Point(105, 109);
            this.lb_Password.Name = "lb_Password";
            this.lb_Password.Size = new System.Drawing.Size(91, 14);
            this.lb_Password.TabIndex = 13;
            this.lb_Password.Text = "Segurança Fraca";
            this.lb_Password.Visible = false;
            // 
            // FormCadastroUser
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 421);
            this.Controls.Add(this.bt_CadastroUser_Sair);
            this.Controls.Add(this.bt_CadastroUser_Ok);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormCadastroUser";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Cadastrar";
            this.Load += new System.EventHandler(this.FormCadastroUser_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox tb_Cadastro_Email;
        private System.Windows.Forms.TextBox tb_Cadastro_Password;
        private System.Windows.Forms.TextBox tb_Cadastro_User;
        private System.Windows.Forms.Button bt_CadastroUser_Ok;
        private System.Windows.Forms.Button bt_CadastroUser_Sair;
        private System.Windows.Forms.ComboBox cb_Cadastro_Equipe;
        private System.Windows.Forms.TextBox tb_Cadastro_Nome;
        private System.Windows.Forms.PictureBox pictureBox5;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Label lb_UserExistente;
        private System.Windows.Forms.Label lb_NomeUserExistente;
        private System.Windows.Forms.Label lb_Password;
    }
}