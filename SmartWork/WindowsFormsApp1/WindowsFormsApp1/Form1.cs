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

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Form2 f2 = new Form2();

        // define the destination folder for button "Open"
        string destinationFolder = " ";

        private string fileName;

        public Form1()
        {
            InitializeComponent();
        }

        //private void panel1_Paint(object sender, PaintEventArgs e)
        //{

        //}
        
        //Button "Add"
        private void button1_Click(object sender, EventArgs e)
        {
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
        }

        //Button "change"
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
        }


        //Button "Open"
        private void button2_Click(object sender, EventArgs e)
        {
            string fname = " ";
            fname = fileName;

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = rowCount;

            for (int i = 1; i <= colCount; i++)
            {
                dataGridView1.Columns[i - 1].HeaderText = xlRange.Cells[1, i].Value2.ToString();
            }
            textBox2.Text = xlRange.Cells[1, 1].Value2;
            textBox3.Text = xlRange.Cells[1, 2].Value2;
            textBox4.Text = xlRange.Cells[1, 3].Value2;
            textBox5.Text = xlRange.Cells[1, 4].Value2;
            textBox6.Text = xlRange.Cells[1, 5].Value2;
            textBox7.Text = xlRange.Cells[1, 6].Value2;
            textBox8.Text = xlRange.Cells[1, 7].Value2;
            textBox9.Text = xlRange.Cells[1, 8].Value2;

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
                textBox10.Text = row.Cells[0].Value.ToString();
                textBox11.Text = row.Cells[1].Value.ToString();
                textBox12.Text = row.Cells[2].Value.ToString();
                textBox13.Text = row.Cells[3].Value.ToString();
                textBox14.Text = row.Cells[4].Value.ToString();
                textBox15.Text = row.Cells[5].Value.ToString();
                textBox16.Text = row.Cells[6].Value.ToString();
                //textBox17.Text = row.Cells[7].Value.ToString();
               
            }
        }
    }
}
