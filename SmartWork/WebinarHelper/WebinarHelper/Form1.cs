using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace WebinarHelper
{
    public partial class Form1 : Form
    {
        Form2 f2 = new Form2();
        Form3 f3 = new Form3();
        Form4 f4 = new Form4();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToLongTimeString();
            label2.Text = DateTime.Now.ToLongDateString();
        }

        // Time Display
        private void timer1_Tick(object sender, EventArgs e)
        {
            label1.Text = DateTime.Now.ToLongTimeString();
            timer1.Start();
        }
        

        private void button1_Click(object sender, EventArgs e)
        {
            f2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = f2.GetPath() + "\\" + textBox1.Text + ".txt";
            //using (File.Create(path));
            richTextBox1.SaveFile(path, RichTextBoxStreamType.PlainText);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string fileLocation = f2.GetPath();
            Process.Start(fileLocation);
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string fileLocation = "C:\\Users\\WB547147\\Documents\\Work Template\\DryrunTemplate";
            Process.Start(fileLocation);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            f3.Show();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            f4.Show();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
