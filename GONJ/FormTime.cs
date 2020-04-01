using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GONJ
{
    public partial class FormTime : Form
    {
        public FormTime()
        {
            InitializeComponent();
        }

        //gavdcodebegin 05
        private void FormTime_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
        }
        //gavdcodeend 05
    }
}
