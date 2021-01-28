using System;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

namespace MCCSI
{
    public partial class Form_ChangePassword : Form
    {
        public Form_ChangePassword()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//submit
        {
            DialogResult result = MessageBox.Show("This will make " + textBox1.Text + " the primary password to login to MCC moving forward. Is this ok?", "Overwrite Previous Password", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                Form1.MCCPassword = textBox1.Text;

                string dataFile = File.ReadAllText(Form1.HiddenFile);
                string[] dataArray = dataFile.Split('|');

                dataArray[3] = textBox1.Text;
                string newString = dataArray[0] + "|" + dataArray[1] + "|" + dataArray[2] + "|" + dataArray[3];

                FileInfo hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Normal;

                File.WriteAllText(Form1.HiddenFile, newString);

                hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Hidden;

                Process.Start(@"C:\Users\14025\source\repos\MCCSI\MCCSI\bin\Debug\MCCTI.exe");
                Application.Exit();
            }
            else if (result == DialogResult.No)
            {
                this.Dispose();
            }
        }

        private void button2_Click(object sender, EventArgs e)//cancel
        {
            this.Dispose();
        }
    }
}
