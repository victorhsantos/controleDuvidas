namespace controleDuvidas
{
    partial class form_AlterarSenha
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form_AlterarSenha));
            this.label1 = new System.Windows.Forms.Label();
            this.tb_Usuario = new System.Windows.Forms.TextBox();
            this.tb_SenhaAtual = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.tb_NovaSenha = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.bt_Salvar = new System.Windows.Forms.Button();
            this.bt_Cancelar = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(60, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Usuário:";
            // 
            // tb_Usuario
            // 
            this.tb_Usuario.BackColor = System.Drawing.SystemColors.Info;
            this.tb_Usuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Usuario.Location = new System.Drawing.Point(26, 39);
            this.tb_Usuario.Name = "tb_Usuario";
            this.tb_Usuario.ReadOnly = true;
            this.tb_Usuario.Size = new System.Drawing.Size(279, 23);
            this.tb_Usuario.TabIndex = 5;
            this.tb_Usuario.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // tb_SenhaAtual
            // 
            this.tb_SenhaAtual.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_SenhaAtual.Location = new System.Drawing.Point(26, 95);
            this.tb_SenhaAtual.MaxLength = 15;
            this.tb_SenhaAtual.Name = "tb_SenhaAtual";
            this.tb_SenhaAtual.PasswordChar = '•';
            this.tb_SenhaAtual.Size = new System.Drawing.Size(279, 23);
            this.tb_SenhaAtual.TabIndex = 1;
            this.tb_SenhaAtual.UseSystemPasswordChar = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 76);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 16);
            this.label2.TabIndex = 2;
            this.label2.Text = "Senha Atual:";
            // 
            // tb_NovaSenha
            // 
            this.tb_NovaSenha.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_NovaSenha.Location = new System.Drawing.Point(26, 153);
            this.tb_NovaSenha.MaxLength = 15;
            this.tb_NovaSenha.Name = "tb_NovaSenha";
            this.tb_NovaSenha.PasswordChar = '•';
            this.tb_NovaSenha.Size = new System.Drawing.Size(279, 23);
            this.tb_NovaSenha.TabIndex = 2;
            this.tb_NovaSenha.UseSystemPasswordChar = true;
            this.tb_NovaSenha.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tb_NovaSenha_KeyPress);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(23, 134);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 16);
            this.label3.TabIndex = 4;
            this.label3.Text = "Nova Senha:";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.tb_NovaSenha);
            this.panel1.Controls.Add(this.tb_Usuario);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.tb_SenhaAtual);
            this.panel1.Location = new System.Drawing.Point(1, 48);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(330, 208);
            this.panel1.TabIndex = 6;
            // 
            // bt_Salvar
            // 
            this.bt_Salvar.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_Salvar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Salvar.Location = new System.Drawing.Point(167, 262);
            this.bt_Salvar.Name = "bt_Salvar";
            this.bt_Salvar.Size = new System.Drawing.Size(88, 39);
            this.bt_Salvar.TabIndex = 4;
            this.bt_Salvar.Text = "Salvar";
            this.bt_Salvar.UseVisualStyleBackColor = false;
            this.bt_Salvar.Click += new System.EventHandler(this.bt_Salvar_Click);
            // 
            // bt_Cancelar
            // 
            this.bt_Cancelar.BackColor = System.Drawing.Color.IndianRed;
            this.bt_Cancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Cancelar.Location = new System.Drawing.Point(73, 262);
            this.bt_Cancelar.Name = "bt_Cancelar";
            this.bt_Cancelar.Size = new System.Drawing.Size(88, 39);
            this.bt_Cancelar.TabIndex = 3;
            this.bt_Cancelar.Text = "Cancelar";
            this.bt_Cancelar.UseVisualStyleBackColor = false;
            this.bt_Cancelar.Click += new System.EventHandler(this.bt_Cancelar_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 15F);
            this.label4.Location = new System.Drawing.Point(104, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(115, 23);
            this.label4.TabIndex = 6;
            this.label4.Text = "Nova Senha";
            // 
            // form_AlterarSenha
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(331, 322);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.bt_Cancelar);
            this.Controls.Add(this.bt_Salvar);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 10F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_AlterarSenha";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Alterar Senha";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_Usuario;
        private System.Windows.Forms.TextBox tb_SenhaAtual;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tb_NovaSenha;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button bt_Salvar;
        private System.Windows.Forms.Button bt_Cancelar;
        private System.Windows.Forms.Label label4;
    }
}