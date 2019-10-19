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

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        string[] fileList;
        
        string destinationFolder = " ";

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            listBox1.AllowDrop = true;
            destinationFolder = textBox1.Text;

            this.listBox1.DragDrop += new
            System.Windows.Forms.DragEventHandler(this.ListBox1_DragDrop);
            this.listBox1.DragEnter += new
            System.Windows.Forms.DragEventHandler(this.ListBox1_DragEnter);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void Form2_DragEnter(object sender, DragEventArgs e)
        {
            
        }

        private void ListBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void ListBox1_DragDrop(object sender, DragEventArgs e)
        {
            fileList = (string[])e.Data.GetData(DataFormats.FileDrop, false);            
            foreach(string file in fileList)
            {
                string filename = getFileName(file);
                MessageBox.Show("You dropped" + file);
                listBox1.Items.Add(filename);
            }
        }

        private void SaveFiles(string[] fileslist)
        {

        }

        private string getFileName(string path)
        {
            return Path.GetFileNameWithoutExtension(path);
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ChooseFolder();
        }

        public void ChooseFolder() {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                destinationFolder = textBox1.Text;
            }         
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //for(int i = 0; i < fileList.Length; i++)
            //{
            //    string sourceFile = fileList[i];
            //    string filename = getFileName(sourceFile);
            //    destinationFolder = destinationFolder + "\\" + filename;
            //    System.IO.File.Move(sourceFile, destinationFolder);
            //}

            foreach (string file in fileList)
            {
                string helper = destinationFolder;
                string filename = getFileName(file);
                string destinationFile = destinationFolder + "\\" + filename + ".xlsx";
                System.IO.File.Move(file, destinationFile);
                destinationFolder = helper;
            }
        }

        public string DestinationFile()
        {
            return textBox1.Text;
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }
    }
}
