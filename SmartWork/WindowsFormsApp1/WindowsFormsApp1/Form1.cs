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
        string destinationFolder = " ";

        private string fileName;

        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            f2.Show();
        }

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

            for (int i = 1; i <= rowCount; i++)
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            fileName = "C:\\Users\\WB547147\\Documents\\Sheila\\" + comboBox1.Text;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                destinationFolder = textBox1.Text;
            }
        }
    }
}
