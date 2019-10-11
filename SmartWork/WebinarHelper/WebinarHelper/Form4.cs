﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;


namespace WebinarHelper
{
    public partial class Form4 : Form
    {
        private string fileName;

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            
        }

 

        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
           
        }

       
        private void button1_Click(object sender, EventArgs e)
        {
            string fname = " ";
            //OpenFileDialog fdlg = new OpenFileDialog();
            //fdlg.Title = "Excel File Dialog";
            //fdlg.InitialDirectory = @"c:\";
            //fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
            //fdlg.FilterIndex = 2;
            //fdlg.RestoreDirectory = true;
            //if (fdlg.ShowDialog() == DialogResult.OK)
            //{
            //    fname = fdlg.FileName;
            //}
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

        private void button2_Click(object sender, EventArgs e)
        {
            Directory.CreateDirectory(@"C:\\Users\\WB547147\\Documents\\Sheila");

            DirectoryInfo dir = new DirectoryInfo(@"C:\\Users\\WB547147\\Documents\\Sheila");
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
    }
}
