namespace controleDuvidas
{
    partial class form_RecuperarSenha
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form_RecuperarSenha));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.tb_Usuario = new System.Windows.Forms.TextBox();
            this.tb_Email = new System.Windows.Forms.TextBox();
            this.bt_Recuperar = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel1.Controls.Add(this.tb_Email);
            this.panel1.Controls.Add(this.tb_Usuario);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Location = new System.Drawing.Point(0, 65);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(362, 171);
            this.panel1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 15F);
            this.label1.Location = new System.Drawing.Point(96, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(163, 23);
            this.label1.TabIndex = 1;
            this.label1.Text = "Recuperar Senha";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(24, 31);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(104, 16);
            this.label2.TabIndex = 0;
            this.label2.Text = "Digite o Usuário:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 91);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 16);
            this.label3.TabIndex = 1;
            this.label3.Text = "Digite o Email:";
            // 
            // tb_Usuario
            // 
            this.tb_Usuario.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Usuario.Location = new System.Drawing.Point(27, 50);
            this.tb_Usuario.MaxLength = 54;
            this.tb_Usuario.Name = "tb_Usuario";
            this.tb_Usuario.Size = new System.Drawing.Size(307, 22);
            this.tb_Usuario.TabIndex = 1;
            // 
            // tb_Email
            // 
            this.tb_Email.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.tb_Email.Location = new System.Drawing.Point(27, 110);
            this.tb_Email.MaxLength = 100;
            this.tb_Email.Name = "tb_Email";
            this.tb_Email.Size = new System.Drawing.Size(307, 22);
            this.tb_Email.TabIndex = 2;
            // 
            // bt_Recuperar
            // 
            this.bt_Recuperar.BackColor = System.Drawing.Color.Khaki;
            this.bt_Recuperar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Recuperar.Location = new System.Drawing.Point(81, 242);
            this.bt_Recuperar.Name = "bt_Recuperar";
            this.bt_Recuperar.Size = new System.Drawing.Size(190, 35);
            this.bt_Recuperar.TabIndex = 3;
            this.bt_Recuperar.Text = "Recuperar";
            this.bt_Recuperar.UseVisualStyleBackColor = false;
            this.bt_Recuperar.Click += new System.EventHandler(this.bt_Recuperar_Click);
            // 
            // form_RecuperarSenha
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(363, 304);
            this.Controls.Add(this.bt_Recuperar);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_RecuperarSenha";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Recuperar Senha";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox tb_Email;
        private System.Windows.Forms.TextBox tb_Usuario;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bt_Recuperar;
    }
}