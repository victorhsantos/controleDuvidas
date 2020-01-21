namespace controleOcorrencias
{
    partial class form_AddAcompanharPRJ
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form_AddAcompanharPRJ));
            this.label1 = new System.Windows.Forms.Label();
            this.clb_Usuarios = new System.Windows.Forms.CheckedListBox();
            this.bt_Adicionar = new System.Windows.Forms.Button();
            this.bt_Cancelar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 10F);
            this.label1.Location = new System.Drawing.Point(26, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(224, 32);
            this.label1.TabIndex = 0;
            this.label1.Text = "Selecione abaixo os usuários que \r\nacompanharam o projeto.";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // clb_Usuarios
            // 
            this.clb_Usuarios.BackColor = System.Drawing.SystemColors.Control;
            this.clb_Usuarios.CheckOnClick = true;
            this.clb_Usuarios.FormattingEnabled = true;
            this.clb_Usuarios.Location = new System.Drawing.Point(46, 59);
            this.clb_Usuarios.Name = "clb_Usuarios";
            this.clb_Usuarios.Size = new System.Drawing.Size(197, 310);
            this.clb_Usuarios.TabIndex = 1;
            // 
            // bt_Adicionar
            // 
            this.bt_Adicionar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.Green;
            this.bt_Adicionar.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.bt_Adicionar.Location = new System.Drawing.Point(146, 375);
            this.bt_Adicionar.Name = "bt_Adicionar";
            this.bt_Adicionar.Size = new System.Drawing.Size(97, 25);
            this.bt_Adicionar.TabIndex = 2;
            this.bt_Adicionar.Text = "Concluído";
            this.bt_Adicionar.UseVisualStyleBackColor = true;
            this.bt_Adicionar.Click += new System.EventHandler(this.bt_Adicionar_Click);
            // 
            // bt_Cancelar
            // 
            this.bt_Cancelar.FlatAppearance.MouseOverBackColor = System.Drawing.Color.IndianRed;
            this.bt_Cancelar.FlatStyle = System.Windows.Forms.FlatStyle.System;
            this.bt_Cancelar.Location = new System.Drawing.Point(46, 375);
            this.bt_Cancelar.Name = "bt_Cancelar";
            this.bt_Cancelar.Size = new System.Drawing.Size(94, 25);
            this.bt_Cancelar.TabIndex = 3;
            this.bt_Cancelar.Text = "Cancelar";
            this.bt_Cancelar.UseVisualStyleBackColor = true;
            this.bt_Cancelar.Click += new System.EventHandler(this.bt_Cancelar_Click);
            // 
            // form_AddAcompanharPRJ
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(280, 412);
            this.Controls.Add(this.bt_Cancelar);
            this.Controls.Add(this.bt_Adicionar);
            this.Controls.Add(this.clb_Usuarios);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Arial", 10F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_AddAcompanharPRJ";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Adicionar Usuários";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckedListBox clb_Usuarios;
        private System.Windows.Forms.Button bt_Adicionar;
        private System.Windows.Forms.Button bt_Cancelar;
    }
}