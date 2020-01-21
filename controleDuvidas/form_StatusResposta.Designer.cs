namespace controleDuvidas
{
    partial class form_StatusResposta
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
            this.cb_StatusResposta = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.bt_Cancelar = new System.Windows.Forms.Button();
            this.bt_Atualizar = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cb_StatusResposta
            // 
            this.cb_StatusResposta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cb_StatusResposta.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cb_StatusResposta.FormattingEnabled = true;
            this.cb_StatusResposta.Location = new System.Drawing.Point(62, 50);
            this.cb_StatusResposta.Name = "cb_StatusResposta";
            this.cb_StatusResposta.Size = new System.Drawing.Size(278, 24);
            this.cb_StatusResposta.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(59, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(141, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Selecione um Status:";
            // 
            // bt_Cancelar
            // 
            this.bt_Cancelar.BackColor = System.Drawing.Color.IndianRed;
            this.bt_Cancelar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Cancelar.Location = new System.Drawing.Point(62, 93);
            this.bt_Cancelar.Name = "bt_Cancelar";
            this.bt_Cancelar.Size = new System.Drawing.Size(114, 39);
            this.bt_Cancelar.TabIndex = 5;
            this.bt_Cancelar.Text = "Cancelar";
            this.bt_Cancelar.UseVisualStyleBackColor = false;
            this.bt_Cancelar.Click += new System.EventHandler(this.bt_Cancelar_Click);
            // 
            // bt_Atualizar
            // 
            this.bt_Atualizar.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.bt_Atualizar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Atualizar.Location = new System.Drawing.Point(222, 93);
            this.bt_Atualizar.Name = "bt_Atualizar";
            this.bt_Atualizar.Size = new System.Drawing.Size(118, 39);
            this.bt_Atualizar.TabIndex = 4;
            this.bt_Atualizar.Text = "Atualizar";
            this.bt_Atualizar.UseVisualStyleBackColor = false;
            this.bt_Atualizar.Click += new System.EventHandler(this.bt_Atualizar_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.cb_StatusResposta);
            this.panel1.Controls.Add(this.bt_Atualizar);
            this.panel1.Controls.Add(this.bt_Cancelar);
            this.panel1.Location = new System.Drawing.Point(0, 28);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(388, 150);
            this.panel1.TabIndex = 8;
            // 
            // form_StatusResposta
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(389, 209);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Arial", 10F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_StatusResposta";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Status Resposta";
            this.Load += new System.EventHandler(this.form_StatusResposta_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cb_StatusResposta;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button bt_Cancelar;
        private System.Windows.Forms.Button bt_Atualizar;
        private System.Windows.Forms.Panel panel1;
    }
}