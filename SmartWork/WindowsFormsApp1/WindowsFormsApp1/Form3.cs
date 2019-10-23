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

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
           
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            
            loadRoomBookingInfo();
            
        }

        private void loadRoomBookingInfo()
        {
            
            string fname = "C:\\Users\\WB547147\\Documents\\Sheila\\RoomBooking.xlsx";
            

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            // dt.Column = colCount;  
            dataGridView1.ColumnCount = colCount;
            dataGridView1.RowCount = rowCount;

            for (int j = 1; j <= colCount; j++)
            {
                dataGridView1.Columns[j - 1].HeaderText = xlRange.Cells[1, j].Value2.ToString();
            }

            for (int i = 2; i <= rowCount; i++)
            {
                if (xlRange.Cells[i, 1].Value2 == null)
                {  
                    continue;
                }
                string sfDFmt = "{0:ddMMyyyy}";
                double smsDbl= xlRange.Cells[i, 1].Value2;
                DateTime smsDate;
                smsDate = DateTime.FromOADate(smsDbl);
                string smsStr = String.Format($"{sfDFmt}", smsDate);
                dataGridView1.Rows[i - 1].Cells[0].Value = smsStr;
            }
            

            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 2; j <= colCount; j++)
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
    }
}
