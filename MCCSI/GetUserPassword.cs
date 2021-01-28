using System.IO;

namespace MCCSI
{
    class GetUserPassword
    {
        public string GetUserPasswordMethod()
        {
            //read file to string variable



            string dataFile = File.ReadAllText(Form1.HiddenFile);
            string[] dataArray = dataFile.Split('|');


            string userPassword = dataArray[3];

            return userPassword;

        }
    }
}
