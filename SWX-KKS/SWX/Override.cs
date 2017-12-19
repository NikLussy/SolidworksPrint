using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SWX_KKS.SWX
{
    public partial class Override : KKS.F_Vorlage 
    {
        public bool AskOverride = true;
        public string Text = "Das PDF existiert bereits";
        
        public Override()
        {
            InitializeComponent();
        }

        private void Override_Load(object sender, EventArgs e)
        {
            lblText.Text = Text;
        }

        private void btnOverride_Click(object sender, EventArgs e)
        {
            AskOverride = !cbRemember.Checked;
            this.DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            AskOverride = !cbRemember.Checked;
            this.DialogResult = DialogResult.Cancel;
        }
    }
}
