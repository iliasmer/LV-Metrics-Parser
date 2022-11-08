using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LV_Metrics_Parser
{
    public partial class MainMenu : Form
    {
        public MainMenu()
        {
            InitializeComponent();
        }

        private void importSourceBtn_Click(object sender, EventArgs e)
        {
            ParseByUsageForm f = new ParseByUsageForm();
            f.Show();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ParseByTimologioForm f = new ParseByTimologioForm();
            f.Show();
        }
    }
}
