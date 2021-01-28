using System;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using Microsoft.Office.Interop.Word;
using System.Threading;

namespace MCCSI
{
    /*
     
        date - version - person - change description

        20210110 - 1.001 - BT - added mcc and kronos login option in program as well as at program initiation.
         
        20210110 - 1.002 - BT - added password change option

        20210110 - 1.003 - BT - fixed password change option

        20210113 - 1.004 - BT - added closing prompt for clocking out reminders

        20210113 - 1.005 - BT - added closing prompt for clocking out reminders

        20210113 - 1.006 - BT - added closing prompt for clocking out reminders to the cancel button

        20210113 - 1.006 - BT - added CheckForUpdates class on Form1 activation

        20210113 - 1.007 - BT - fix CheckForUpdates class on Form1 activation using \\localhost\B$\MCCSI in url path in MCCSI project properties

        20210113 - 1.008 - BT - test1 CheckForUpdates class on Form1 activation using \\localhost\B$\MCCSI in url path in MCCSI project properties

        20210113 - 1.009 - BT - fix CheckForUpdates class on Form1 activation using http://LAPTOP-A5V674K0/SampleApplication/ in url path in MCCSI project properties

        20210113 - 1.010 - BT - fix CheckForUpdates class on Form1 activation using //localhost/B$/MCCSI/ in url path in MCCSI project properties

        20210113 - 1.011 - BT - test1 CheckForUpdates class on Form1 activation using //localhost/B$/MCCSI/ in url path in MCCSI project properties

        20210113 - 1.012 - BT - test2 CheckForUpdates class on Form1 activation using //localhost/B$/MCCSI/ in url path in MCCSI project properties

        20210124 - 1.013 - BT - fixed indexing errors related to taking data from the .psv file and Auto-Login option.

        20210124 - 1.014 - BT - changed the name to MCCTI (MCC Tutor Interface).

        20210124 - 1.015 - BT - Re-added student folder creation.

        20210124 - 1.016 - BT - Changed the name in another part of the project.

        20210124 - 1.017 - BT - updated all paths with MCCSI to MCCTI.


         */


    public partial class Form1 : Form
    {
        
        private static string StudentsFolder;
        private static string StudentsFileName;
        private static Word.Table TheTable;
        private string windowsUserName = System.Environment.UserName;//gives windows username
        public static string TutorsName;
        public static string AutoOpenFlag;
        public static string HiddenFile;
        public static string TutorRoom;
        public static string MCCPassword;
        public static string MCCsWebsite = @"https://sso.mccneb.edu/adfs/ls?wa=wsignin1.0&wtrealm=urn%3amyway.mccneb.edu%3a443&wctx=https%3a%2f%2fmyway.mccneb.edu%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d%252F&wreply=https%3a%2f%2fmyway.mccneb.edu%2f_trust%2fdefault.aspx";

        private static GetTutorNameFromFile getTutorName = new GetTutorNameFromFile();
        private static GetCheckBoxFlag getCheckBoxFlag = new GetCheckBoxFlag();
        private static GetMeetingRoomLink getMeetingRoomLink = new GetMeetingRoomLink();
        private static GetUserPassword getUserPassword = new GetUserPassword();



        public Form1()
        {
            InitializeComponent();

            textBox1.TabIndex = 0;
            textBox2.TabIndex = 1;
            textBox4.TabIndex = 3;
            textBox5.TabIndex = 4;
            button1.TabIndex = 5;
            button3.TabIndex = 6;

            this.StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            
            
            
            string hiddenPSVFile = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\miscData.psv";
            string studentsName = textBox1.Text;
            string studentsClass = textBox2.Text;
            

            string problemType = textBox4.Text;
            string numberOfProblems = textBox5.Text;
            

            HiddenFile = hiddenPSVFile;
            
            this.Text = "MCCTI - "+ getTutorName.GetTutorNameFromFileMethod();


            string studentsFolder = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Students\";
            string studentFolder = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Students\" + studentsName + @"\";

            string studentsFileName = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Students\" + studentsName + @"\" + problemType + @" with " + TutorsName + ".docx";
            string mathProblemTemplate = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Tools\Math Problem Template.docx";
            
            
            StudentsFileName = studentsFileName;
            if (!File.Exists(hiddenPSVFile)|| !Directory.Exists(studentsFolder))
            {
                try
                {
                    File.Create(hiddenPSVFile);

                    string setData = "Tutor Name|T|Some Link Here|Some Password";

                    
                    

                    File.WriteAllText(hiddenPSVFile, setData);

                    FileInfo hiddenFile = new FileInfo(hiddenPSVFile);
                    hiddenFile.Attributes = FileAttributes.Hidden;

                    try
                    {
                        Directory.CreateDirectory(studentsFolder);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }


                }
                catch (IOException ex)
                {
                    MessageBox.Show("Could not create file because \r\r" + ex, "Could Not Create File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                TutorsName = getTutorName.GetTutorNameFromFileMethod();

                AutoOpenFlag = getCheckBoxFlag.GetCheckBoxFlagMethod();

                TutorRoom = getMeetingRoomLink.GetMeetingRoomLinkMethod();

                MCCPassword = getUserPassword.GetUserPasswordMethod();

                if (AutoOpenFlag == "T")
                {

                    checkBox1.Checked = true;

                }
                else if (AutoOpenFlag == "F")
                {
                    checkBox1.Checked = false;
                }
            }


           

           

        }

        private void button1_Click(object sender, EventArgs e)//Submit
        {
          

            string studentsName = textBox1.Text;
            string studentsClass = textBox2.Text;
            
            string problemType = textBox4.Text;
            string numberOfProblems = textBox5.Text;

            string hiddenPSVFile = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\miscData.psv";
            HiddenFile = hiddenPSVFile;
            
            string studentFolder = @"C:\Users\"+ windowsUserName + @"\Documents\Math Tutoring\Students\"+studentsName+@"\";
            string studentsFileName = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Students\" + studentsName + @"\"+ problemType+@" with "+ TutorsName + ".docx";
            string mathProblemTemplate = @"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Tools\Math Problem Template.docx";
            StudentsFileName = studentsFileName;
            DateTime now = DateTime.Now;



            Word.Application app = new Word.Application();
            app.WindowState = Word.WdWindowState.wdWindowStateNormal;
            Word.Document doc = app.Documents.Add();
            Word.Table table;

            foreach (Word.Section section in doc.Sections)
            {
                DateTime today = DateTime.Today;
                Word.Range headerRange1 = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //Word.Range headerRange2 = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                headerRange1.Fields.Add(headerRange1, WdFieldType.wdFieldPage);
                headerRange1.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                headerRange1.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange1.Font.Size = 10;
                headerRange1.Text = "MCC FOC Online Math Tutoring"+Environment.NewLine+ today.ToString("MM/dd/yyyy");


                //headerRange2.Fields.Add(headerRange2, WdFieldType.wdFieldPage);
                //headerRange2.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                //headerRange2.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                //headerRange2.Font.Size = 10;
                //headerRange2.Text = today.ToString("MM/dd/yyyy"); 

            }


            app.Visible = true;








            //create students folder


            if (!Directory.Exists(studentFolder))
            {
                try
                {
                    Directory.CreateDirectory(studentFolder);

                    



                }
                catch (IOException ex)
                {
                    MessageBox.Show("Could not create directory for the student because \r\r" + ex, "Could Not Create Directory", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //MessageBox.Show("This student already has a folder", "Previously Existing Student!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }


            if (checkBox1.Checked == true)
            {

                
                Process.Start(studentFolder);

            }
            else if(checkBox1.Checked == false)
            {

            }

            if (File.Exists(mathProblemTemplate))
            {
                Process.Start(mathProblemTemplate);
            }
            
            //object objMiss = System.Reflection.Missing.Value;
            //Word.Range objWordRng = doc.Bookmarks.get_Item(ref objMiss).Range; //go to end of document

            //Create file and insert student data into word file

            if (!File.Exists(studentsFileName))
            {
                //table = app.Selection.Tables.Add(objWordRng, 1, 1);

                string stringToInsert = "Student: " + studentsName + "\r" +
                                        "Class: " + studentsClass + "\r" +
                                        "Problem Type: " + problemType + "\r" +
                                        "Number of Problems: " + numberOfProblems + "\r" +
                                        "Time Spent: From "+now.ToString("hh:mm tt")+" to ";//wanted to add tables here but couldn't get it to work
                                        



                try
                {
                    Word.Paragraph writeToDoc;

                    writeToDoc = doc.Paragraphs.Add();
                    
                    writeToDoc.Range.Text = stringToInsert;
                    
                    //foreach (string strng in stringArray)
                    //{
                    //    doc.Content.Text = stringToInsert;
                    //}

                    doc.SaveAs2(studentsFileName);
                    //doc.Close();
                    //app.Quit();
                    textBox1.Text = "";
                    textBox2.Text = "";
                    textBox5.Text = "";
                    textBox4.Text = "";
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Could not create file for the student because \r\r" + ex, "Could Not File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            //this.Close();  
        }


   

        private void button3_Click(object sender, EventArgs e)//Cancel
        {
            DialogResult dialogResult = MessageBox.Show("Have you clocked out already?", "Clock Out Prompt", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);

            if (dialogResult == DialogResult.Yes)
            {
                this.Dispose();
            }
            else if (dialogResult == DialogResult.No)
            {
                //login into metro's website
                Process.Start("https://sso.mccneb.edu/adfs/ls?wa=wsignin1.0&wtrealm=urn%3amyway.mccneb.edu%3a443&wctx=https%3a%2f%2fmyway.mccneb.edu%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d%252F&wreply=https%3a%2f%2fmyway.mccneb.edu%2f_trust%2fdefault.aspx");

                Thread.Sleep(9000);

                SendKeys.Send("bturner4@mccneb.edu");
                SendKeys.Send("{TAB}");
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                //login into kronos

                Process.Start("https://mccneb.kronos.net/wfc/logon");
                Thread.Sleep(3000);
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                this.Dispose();
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                //do nothing
            }
        }

        private void changeTutorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_GetTutorName changeTutor = new Form_GetTutorName();
            changeTutor.StartPosition = FormStartPosition.CenterScreen;
            changeTutor.Show();
            
        }
        private void changePasswordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ChangePassword changePassword = new Form_ChangePassword();
            changePassword.StartPosition = FormStartPosition.CenterScreen;
            changePassword.Show();
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {

                string dataFile = File.ReadAllText(Form1.HiddenFile);
                string[] dataArray = dataFile.Split('|');

                dataArray[1] = "T";
                string newString = dataArray[0] + "|" + dataArray[1] + "|" + dataArray[2] + "|" + dataArray[3];


                FileInfo hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Normal;

                File.WriteAllText(Form1.HiddenFile, newString);

                hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Hidden;




            }
            else if (checkBox1.Checked == false)
            {
                string dataFile = File.ReadAllText(Form1.HiddenFile);
                string[] dataArray = dataFile.Split('|');

                dataArray[1] = "F";
                string newString = dataArray[0] + "|" + dataArray[1] + "|" + dataArray[2] + "|" + dataArray[3];

                FileInfo hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Normal;

                File.WriteAllText(Form1.HiddenFile, newString);

                hiddenFile = new FileInfo(Form1.HiddenFile);
                hiddenFile.Attributes = FileAttributes.Hidden;



            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Process.Start(@"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Students\");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            DialogResult dialogResult = MessageBox.Show("Have you clocked out already?","Clock Out Prompt",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Exclamation);

            if (dialogResult == DialogResult.Yes)
            {
                //do nothing and continue closing
            }
            else if (dialogResult == DialogResult.No)
            {
                //login into metro's website
                Process.Start("https://sso.mccneb.edu/adfs/ls?wa=wsignin1.0&wtrealm=urn%3amyway.mccneb.edu%3a443&wctx=https%3a%2f%2fmyway.mccneb.edu%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d%252F&wreply=https%3a%2f%2fmyway.mccneb.edu%2f_trust%2fdefault.aspx");

                Thread.Sleep(9000);

                SendKeys.Send("bturner4@mccneb.edu");
                SendKeys.Send("{TAB}");
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                //login into kronos

                Process.Start("https://mccneb.kronos.net/wfc/logon");
                Thread.Sleep(3000);
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                //punch into kronos
            }
            else if (dialogResult == DialogResult.Cancel)
            {
                e.Cancel = true;
            }


        }

        private void openMeetingRoomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string link = Form_ChangeLink.TextBox1;

            if (Form_ChangeLink.ChangedLink == true)
            {

            }
            else
            {
                Process.Start(TutorRoom);

            }

        }

        private void button4_Click(object sender, EventArgs e)//Tools
        {
            Process.Start(@"C:\Users\" + windowsUserName + @"\Documents\Math Tutoring\Tools\");

        }

        private void changeMeetingRoomLinkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form_ChangeLink form = new Form_ChangeLink();
            
            form.Show();
        }

        private void button5_Click(object sender, EventArgs e)//clear fields button
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox5.Text = "";
            textBox4.Text = "";
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Control && e.KeyCode == Keys.T)
            {
                Process.Start(@"C:\Users\14025\Documents\Math Tutoring");
                e.Handled = true;
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Control && e.KeyCode == Keys.T)
            {
                Process.Start(@"C:\Users\14025\Documents\Math Tutoring");
                e.Handled = true;
            }
        }

        private void loginToMCCAndPunchInToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
            //login into metro's website
            Process.Start(MCCsWebsite);

            Thread.Sleep(8000);

            SendKeys.Send("bturner4@mccneb.edu");
            SendKeys.Send("{TAB}");
            SendKeys.Send(MCCPassword);
            SendKeys.Send("{ENTER}");

            //login into kronos

            Process.Start("https://mccneb.kronos.net/wfc/logon");
            Thread.Sleep(3000);
            SendKeys.Send(MCCPassword);
            SendKeys.Send("{ENTER}");

            //punch into kronos



        }

        private void Form1_Shown(object sender, EventArgs e)
        {

            DialogResult result = MessageBox.Show("Do you want to use Auto-login option now?","MCCTI's Auto-Login",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                //login into metro's website
                Process.Start(MCCsWebsite);

                Thread.Sleep(8000);

                SendKeys.Send("bturner4@mccneb.edu");
                SendKeys.Send("{TAB}");
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                //login into kronos

                Process.Start("https://mccneb.kronos.net/wfc/logon");
                Thread.Sleep(3000);
                SendKeys.Send(MCCPassword);
                SendKeys.Send("{ENTER}");

                //punch into kronos
            }
            else if (result == DialogResult.No)
            {
                //Do nothing
            }
            else if (result == DialogResult.Cancel)
            {
                this.Dispose();
            }


        }

        //private void Form1_Activated(object sender, EventArgs e)
        //{
        //    CheckForUpdates checkForUpdates = new CheckForUpdates();
        //    checkForUpdates.InstallUpdateSyncWithInfo();
        //}

        private void checkForUpdatesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CheckForUpdates checkForUpdates = new CheckForUpdates();
            checkForUpdates.InstallUpdateSyncWithInfo();
        }
    }
}
