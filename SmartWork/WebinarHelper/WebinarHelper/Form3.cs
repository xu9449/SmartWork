using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace WebinarHelper
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sName = textBox1.Text;
            comboBox1.Items.Add(sName);
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox2.Text = comboBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string fileLocation = "C:\\Users\\WB547147\\Documents\\Work Template\\DryrunTemplate";
            Process.Start(fileLocation);
        }

        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();

        }
    }
}
