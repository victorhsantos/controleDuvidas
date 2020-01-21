namespace controleOcorrencias
{
    partial class form_AddDataPrev
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(form_AddDataPrev));
            this.calendar_DataPrev = new System.Windows.Forms.MonthCalendar();
            this.bt_Adicionar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // calendar_DataPrev
            // 
            this.calendar_DataPrev.Font = new System.Drawing.Font("Arial", 12F);
            this.calendar_DataPrev.Location = new System.Drawing.Point(18, 18);
            this.calendar_DataPrev.MaxSelectionCount = 1;
            this.calendar_DataPrev.Name = "calendar_DataPrev";
            this.calendar_DataPrev.ShowTodayCircle = false;
            this.calendar_DataPrev.TabIndex = 0;
            // 
            // bt_Adicionar
            // 
            this.bt_Adicionar.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.bt_Adicionar.Location = new System.Drawing.Point(166, 187);
            this.bt_Adicionar.Name = "bt_Adicionar";
            this.bt_Adicionar.Size = new System.Drawing.Size(79, 23);
            this.bt_Adicionar.TabIndex = 1;
            this.bt_Adicionar.Text = "Adicionar";
            this.bt_Adicionar.UseVisualStyleBackColor = true;
            this.bt_Adicionar.Click += new System.EventHandler(this.bt_Adicionar_Click);
            // 
            // form_AddDataPrev
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(265, 221);
            this.Controls.Add(this.bt_Adicionar);
            this.Controls.Add(this.calendar_DataPrev);
            this.Font = new System.Drawing.Font("Arial", 10F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "form_AddDataPrev";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Data Prevista Solução";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.MonthCalendar calendar_DataPrev;
        private System.Windows.Forms.Button bt_Adicionar;
    }
}