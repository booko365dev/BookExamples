﻿using System;
using System.Windows.Forms;

namespace TFVB
{
    public partial class FormTime : Form
    {
        //gavdcodebegin 004
        public FormTime()
        {
            InitializeComponent();
        }

        private void FormTime_Load(object sender, EventArgs e)
        {
            lblTime.Text = DateTime.Now.ToLongTimeString();
        }
        //gavdcodeend 004
    }
}
