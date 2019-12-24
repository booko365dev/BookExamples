using System;
using System.Windows.Forms;

namespace TFVB
{
    public partial class FormTime : Form
    {
        public FormTime()
        {
            InitializeComponent();
        }

        private void FormTime_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
        }
    }
}

