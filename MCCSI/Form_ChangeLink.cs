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

namespace MCCSI
{
    public partial class Form_ChangeLink : Form
    {
        public static string TextBox1;
        public static bool ChangedLink = false;

        
        
        
        public Form_ChangeLink()
        {
            InitializeComponent();
        }

        private void Form_ChangeLink_Load(object sender, EventArgs e)
        {
            this.Location = new Point(500,510);
        }

        private void button1_Click(object sender, EventArgs e)//Submit
        {

            ;
            string newString = Form1.TutorsName + "|" + Form1.AutoOpenFlag + "|" + textBox1.Text +"|" + Form1.MCCPassword;


            FileInfo hiddenFile = new FileInfo(Form1.HiddenFile);
            hiddenFile.Attributes = FileAttributes.Normal;

            File.WriteAllText(Form1.HiddenFile, newString);

            hiddenFile = new FileInfo(Form1.HiddenFile);
            hiddenFile.Attributes = FileAttributes.Hidden;

            Process.Start(@"C:\Users\14025\source\repos\MCCSI\MCCSI\bin\Debug\MCCTI.exe");
            Application.Exit();
            this.Dispose();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}
