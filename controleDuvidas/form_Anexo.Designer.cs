namespace controleOcorrencias
{
    partial class form_Anexo
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form_Anexo));
            this.label1 = new System.Windows.Forms.Label();
            this.tb_LocalArquivo = new System.Windows.Forms.TextBox();
            this.bt_Abrir = new System.Windows.Forms.Button();
            this.bt_RemoverAnexo = new System.Windows.Forms.Button();
            this.bt_Cancelar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Local Arquivo:";
            // 
            // tb_LocalArquivo
            // 
            this.tb_LocalArquivo.Location = new System.Drawing.Point(115, 49);
            this.tb_LocalArquivo.Name = "tb_LocalArquivo";
            this.tb_LocalArquivo.ReadOnly = true;
            this.tb_LocalArquivo.Size = new System.Drawing.Size(329, 23);
            this.tb_LocalArquivo.TabIndex = 1;
            // 
            // bt_Abrir
            // 
            this.bt_Abrir.Location = new System.Drawing.Point(357, 90);
            this.bt_Abrir.Name = "bt_Abrir";
            this.bt_Abrir.Size = new System.Drawing.Size(87, 23);
            this.bt_Abrir.TabIndex = 2;
            this.bt_Abrir.Text = "Abrir";
            this.bt_Abrir.UseVisualStyleBackColor = true;
            this.bt_Abrir.Click += new System.EventHandler(this.bt_Abrir_Click);
            // 
            // bt_RemoverAnexo
            // 
            this.bt_RemoverAnexo.Location = new System.Drawing.Point(264, 90);
            this.bt_RemoverAnexo.Name = "bt_RemoverAnexo";
            this.bt_RemoverAnexo.Size = new System.Drawing.Size(87, 23);
            this.bt_RemoverAnexo.TabIndex = 3;
            this.bt_RemoverAnexo.Text = "Remover";
            this.bt_RemoverAnexo.UseVisualStyleBackColor = true;
            this.bt_RemoverAnexo.Click += new System.EventHandler(this.bt_RemoverAnexo_Click);
            // 
            // bt_Cancelar
            // 
            this.bt_Cancelar.Location = new System.Drawing.Point(171, 90);
            this.bt_Cancelar.Name = "bt_Cancelar";
            this.bt_Cancelar.Size = new System.Drawing.Size(87, 23);
            this.bt_Cancelar.TabIndex = 4;
            this.bt_Cancelar.Text = "Cancelar";
            this.bt_Cancelar.UseVisualStyleBackColor = true;
            this.bt_Cancelar.Click += new System.EventHandler(this.bt_Cancelar_Click);
            // 
            // form_Anexo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(456, 141);
            this.Controls.Add(this.bt_Cancelar);
            this.Controls.Add(this.bt_RemoverAnexo);
            this.Controls.Add(this.bt_Abrir);
            this.Controls.Add(this.tb_LocalArquivo);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Arial", 10F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_Anexo";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Anexo";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox tb_LocalArquivo;
        private System.Windows.Forms.Button bt_Abrir;
        private System.Windows.Forms.Button bt_RemoverAnexo;
        private System.Windows.Forms.Button bt_Cancelar;
    }
}