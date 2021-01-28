using System;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

namespace MCCSI
{
    public partial class Form_GetTutorName : Form
    {

        Form1 form1 = new Form1();



        public Form_GetTutorName()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.Manual;
        }

        private void button1_Click(object sender, EventArgs e)//submit
        {

            DialogResult result = MessageBox.Show("This will make "+textBox1.Text+" the primary user of this app. All files will be saved with this name as the Math tutor. Is this ok?","Overwrite Previous User",MessageBoxButtons.YesNo,MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                
                string newString = textBox1.Text + "|" + Form1.AutoOpenFlag + "|" + Form1.TutorRoom +"|"+ Form1.MCCPassword;

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
            //Process.Start(@"C:\Users\14025\source\repos\MCCSI\MCCSI\bin\Debug\MCCSI.exe");
            //Application.Exit();
        }

      
    }
}
