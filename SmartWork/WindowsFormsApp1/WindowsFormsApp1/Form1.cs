using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;
using System.Data.OleDb;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Form2 f2 = new Form2();
        Form3 f3 = new Form3();

        // define the destination folder for button "Open"
        string destinationFolder = " ";

        private string fileName;

        public Form1()
        {
            InitializeComponent();
        }
        
        //Button "Add"
        private void button1_Click(object sender, EventArgs e)
        {
            if(f2.IsDisposed)
            {
                f2 = new Form2();
                f2.Show();
            }
            f2.Show();
        }

        //Button "Load"
        private void button3_Click(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            Directory.CreateDirectory(textBox1.Text);
            DirectoryInfo dir = new DirectoryInfo(textBox1.Text);
            FileInfo[] files = dir.GetFiles("*.xlsx");
            foreach (FileInfo file in files)
            {
                comboBox1.Items.Add(file);
            }
            destinationFolder = textBox1.Text;
        }

        //Button "Choose"
        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                destinationFolder = textBox1.Text;
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            fileName = destinationFolder +"\\" + comboBox1.Text;
            showFiles(1);
        }


        //Button "Open"
        private void button2_Click(object sender, EventArgs e)
        {
            showFiles(1);
        }

        private void showFiles(int page)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            string fname = " ";
            fname = fileName;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[page];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = rowCount;

            for (int i = 1; i <= colCount; i++)
            {
                if (xlRange.Cells[1, i].Value2 == null)
                {
                    xlRange.Cells[1, i].Value2 = " ";
                }
                dataGridView1.Columns[i - 1].HeaderText = xlRange.Cells[1, i].Value2.ToString();
            }


            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    //write the value to the Grid  
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                    }
                    // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

                    //add useful things here!     
                }
            }

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Application.Exit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                int x = dataGridView1.ColumnCount;
                for (int i = 0; i< x; i++)
                {
                    if (row.Cells[i].Value == null)
                        row.Cells[i].Value = " - ";
                }
                
               
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            f3.Show();
        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {
            showFiles(2);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            showFiles(3);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            showFiles(4);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            showFiles(5);
        }
    }
}
