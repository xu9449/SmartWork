using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;

namespace WebinarHelper
{
    using Word = Microsoft.Office.Interop.Word;

    public partial class Form2 : Form
    {
        string[] fileList;
        string destinationFile = " ";
        string destinationFolder = " ";

        public Form2()
        {
            InitializeComponent();
            comboBox1.Items.Add("speaker");      
    }
    

        private void Form2_Load(object sender, EventArgs e)
        {
            richTextBox1.AllowDrop = true;
            richTextBox1.DragDrop += RichTextBox1_DragDrop;
        }

        private void RichTextBox1_DragDrop(object sender, DragEventArgs e)
        {
            fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            ImportWord(fileList);
        }

        public string GetPath()
        {
            return destinationFolder;
        }
        

        void ImportWord(string[] fileList)
        {
            Microsoft.Office.Interop.Word.Application wordObject = new Microsoft.Office.Interop.Word.Application();
            object File = fileList[0]; //this is the path
            object nullobject = System.Reflection.Missing.Value; Microsoft.Office.Interop.Word.Application wordobject = new Microsoft.Office.Interop.Word.Application();
            wordobject.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone; Microsoft.Office.Interop.Word._Document docs = wordObject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject); docs.ActiveWindow.Selection.WholeStory();
            docs.ActiveWindow.Selection.Copy();
            this.richTextBox1.Paste();
            docs.Close(ref nullobject, ref nullobject, ref nullobject);
            wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
        }



        private void button1_Click(object sender, EventArgs e)
        {
            string date = this.dateTimePicker1.Text;
            date = new string((from c in date
                               where char.IsLetterOrDigit(c)
                               select c
                                ).ToArray());
            
            string path = textBox2.Text + "\\" + date + textBox1.Text;
            destinationFile = path + "\\RequestForm.docx";
            destinationFolder = path;
            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                MessageBox.Show("Directory Created!");
            } else
            {
                MessageBox.Show("Directory Already Exists.");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start("chrome", @"https://worldbankva.adobeconnect.com/admin/meeting/folder/list?filter-rows=100&filter-start=0&sco-id=1316836921&tab-id=833642799&sort-date-begin=desc&OWASP_CSRFTOKEN=c00503813114f1ea897c074045b9d008c3977579a450d73c893d0a0eb5f43a51");
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Process.Start("chrome", @"https://worldbankva.adobeconnect.com/admin/event/folder/list?filter-rows=100&filter-start=0&sco-id=1088621214&tab-id=833642803&sort-date-begin=desc&OWASP_CSRFTOKEN=c00503813114f1ea897c074045b9d008c3977579a450d73c893d0a0eb5f43a51");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Process.Start(fileList[0]);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string sourceFile = fileList[0];

            System.IO.File.Move(sourceFile, destinationFile);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Process.Start("chrome", @"https://olc.worldbank.org/staff-learning");
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //String temp = richTextBox1.Text;
            //richTextBox1.Text = " ";
            //richTextBox1.Text = temp;
            int index = 0;
            while (index < richTextBox1.Text.LastIndexOf(comboBox1.Text))
            {
                richTextBox1.Find(comboBox1.Text, index, richTextBox1.TextLength, RichTextBoxFinds.None);
                richTextBox1.SelectionBackColor = Color.Yellow;
                index = richTextBox1.Text.IndexOf(comboBox1.Text, index) + 1;
            }
        }
    }
}
