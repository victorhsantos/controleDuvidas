using System;
using System.Windows.Forms;
using System.Reflection;

namespace controleOcorrencias
{
    public partial class formInfo : Form
    {
        public formInfo()
        {
            InitializeComponent();
        }

        private void formInfo_Load(object sender, EventArgs e)
        {
            Version versao = Assembly.GetExecutingAssembly().GetName().Version;
            this.lb_versao.Text = "Versão " + versao.ToString().Substring(0, 3);
        }
    }
}
