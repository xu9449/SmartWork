﻿using System;
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
            
                }

        private void Form3_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }
    }

            
        }
