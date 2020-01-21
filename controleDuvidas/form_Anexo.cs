using System;
using System.IO;
using System.Windows.Forms;
using controleDuvidas;

namespace controleOcorrencias
{
    public partial class form_Anexo : Form
    {       
        public form_Anexo(string caminhoOrigem)
        {
            InitializeComponent();

            tb_LocalArquivo.Text = caminhoOrigem;            
        }

        private void bt_Abrir_Click(object sender, EventArgs e)
        {
            if (File.Exists(tb_LocalArquivo.Text))
                System.Diagnostics.Process.Start(tb_LocalArquivo.Text);
        }

        private void bt_RemoverAnexo_Click(object sender, EventArgs e)
        {
            Form1 f1 = new Form1();
            f1.removeAnexo = 1;            
            this.Close();
        }

        private void bt_Cancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
